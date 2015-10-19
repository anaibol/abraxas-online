Attribute VB_Name = "GeneralES"

Option Explicit
'********************************************************************************
'********************************************************************************
'********************************************************************************
'*********************** Funciones de Carga *************************************
'********************************************************************************
'********************************************************************************
Public grhCount As Long
Public Sub LoadGrhData(Optional ByVal FileNamePath As String = vbNullString)
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Long
Dim Frame As Long
Dim TempInt As Integer
Dim ArchivoAbrir As String
Dim fileVersion As Long

'Open files

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = IniPath & "Grh.ind"
    Else
        ArchivoAbrir = IniPath & "Graficos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If

Open ArchivoAbrir For Binary Access Read As #1
Seek #1, 1

Get #1, , fileVersion
Get #1, , grhCount

ReDim GrhData(1 To MAXGrH) As GrhData

While Not EOF(1)

    'Fill Grh List
    Get #1, , Grh
    If Grh = 0 Then GoTo Despues
    With GrhData(Grh)
        'Get number of frames
        Get #1, , .NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        ReDim .Frames(1 To GrhData(Grh).NumFrames)
        
        If GrhData(Grh).NumFrames > 1 Then
            frmMain.Lista.AddItem Grh & " (animacion)"
            
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #1, , .Frames(Frame)
                If .Frames(Frame) <= 0 Or .Frames(Frame) > MAXGrH Then
                    GoTo ErrorHandler
                End If
            Next Frame
        
            Get #1, , .speed
            If .speed <= 0 Then MsgBox Grh & " velocidad <= 0 ", , "advertencia"
            
            'Compute width and height
            .pixelHeight = GrhData(.Frames(1)).pixelHeight
            If .pixelHeight <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
            
            .pixelWidth = GrhData(.Frames(1)).pixelWidth
            If .pixelWidth <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
            
            .TileWidth = GrhData(.Frames(1)).TileWidth
            If .TileWidth <= 0 Then MsgBox Grh & " anchoT <= 0 ", , "advertencia"
            
            .TileHeight = GrhData(.Frames(1)).TileHeight
            If .TileHeight <= 0 Then MsgBox Grh & " altoT <= 0 ", , "advertencia"
        
        Else
            frmMain.Lista.AddItem Grh
            'Read in normal GRH data
            Get #1, , .FileNum
            If .FileNum <= 0 Then MsgBox Grh & " tiene bmp = 0 ", , "advertencia"
               
            Get #1, , .sX
            If .sX < 0 Then MsgBox Grh & " tiene Sx <= 0 ", , "advertencia"
            
            Get #1, , .sY
            If .sY < 0 Then MsgBox Grh & " tiene Sy <= 0 ", , "advertencia"
                
            Get #1, , .pixelWidth
            If .pixelWidth <= 0 Then MsgBox Grh & " ancho <= 0 ", , "advertencia"
            
            Get #1, , .pixelHeight
            If .pixelHeight <= 0 Then MsgBox Grh & " alto <= 0 ", , "advertencia"
            
            'Compute width and height
            .TileWidth = .pixelWidth / TilePixelHeight
            .TileHeight = .pixelHeight / TilePixelWidth
            
            GrhData(Grh).Frames(1) = Grh
        End If
    End With
Wend
'************************************************

Despues:

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub
Public Function LoadGrhData2() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Integer
    Dim handle As Integer
    Dim TempInt As Integer
    Dim tmpLng  As Long
    
    'Open files
    handle = FreeFile()
    Open App.Path & "\Graficosv2.ind" For Binary Access Read As handle
    
        Seek handle, 1
        Get handle, , tmpLng
        Get handle, , tmpLng
        While Not EOF(handle)
            Get handle, , Grh
    
              With GrhData2(Grh)
                  'Get number of frames
                  Get handle, , .NumFrames
                  If .NumFrames <= 0 Then GoTo ErrorHandler
                  
                  ReDim .Frames(1 To .NumFrames)
                  
                  If .NumFrames > 1 Then
                      'Read a animation GRH set
                      For Frame = 1 To .NumFrames
                          Get handle, , .Frames(Frame)
                          If .Frames(Frame) <= 0 Then
                              GoTo ErrorHandler
                          End If
                      Next Frame

                      Get handle, , .speed
                      If .speed <= 0 Then GoTo ErrorHandler
                      
                      'Compute width and height
                      .pixelHeight = GrhData2(.Frames(1)).pixelHeight
                      If .pixelHeight <= 0 Then GoTo ErrorHandler
                      
                      .pixelWidth = GrhData2(.Frames(1)).pixelWidth
                      If .pixelWidth <= 0 Then GoTo ErrorHandler
                      
                      .TileWidth = GrhData2(.Frames(1)).TileWidth
                      If .TileWidth <= 0 Then GoTo ErrorHandler
                      
                      .TileHeight = GrhData2(.Frames(1)).TileHeight
                      If .TileHeight <= 0 Then GoTo ErrorHandler
                  Else
                      
                      'Read in normal GRH data
                      Get handle, , .FileNum
                      If .FileNum <= 0 Then GoTo ErrorHandler
                      
                      Get handle, , .sX
                      If .sX < 0 Then GoTo ErrorHandler
                      
                      Get handle, , .sY
                      If .sY < 0 Then GoTo ErrorHandler
                      
                      Get handle, , .pixelWidth
                      If .pixelWidth <= 0 Then GoTo ErrorHandler
                      
                      Get handle, , .pixelHeight
                      If .pixelHeight <= 0 Then GoTo ErrorHandler
                      
                      'Compute width and height
                      .TileWidth = .pixelWidth / 32
                      .TileHeight = .pixelHeight / 32
                      
                      .Frames(1) = Grh
                  End If
              End With
        Wend
    Close handle
     
    LoadGrhData2 = True
Exit Function

ErrorHandler:
    LoadGrhData2 = False
    Close handle
End Function

Public Sub CargarCuerposDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer

Dim MisCuerpos() As tIndiceCuerpoLong
Dim ArchivoAbrir As String
Dim loopc As Integer


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Body.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"

    Exit Sub
End If
Dim Leer As New clsIniReader

Call Leer.Initialize(ArchivoAbrir)

NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpoLong

ReDim BodyData(0 To NumCuerpos) As BodyData

For loopc = 1 To NumCuerpos
    InitGrh BodyData(loopc).Walk(1), Val(Leer.GetValue("Body" & loopc, "WALK1")), 0
    InitGrh BodyData(loopc).Walk(2), Val(Leer.GetValue("Body" & loopc, "WALK2")), 0
    InitGrh BodyData(loopc).Walk(3), Val(Leer.GetValue("Body" & loopc, "WALK3")), 0
    InitGrh BodyData(loopc).Walk(4), Val(Leer.GetValue("body" & loopc, "WALK4")), 0
    BodyData(loopc).HeadOffset.X = Val(Leer.GetValue("body" & loopc, "HeadOffsetX"))
    BodyData(loopc).HeadOffset.Y = Val(Leer.GetValue("body" & loopc, "HeadOffsety"))
Next loopc
Set Leer = Nothing

End Sub


Public Sub CargarCabezasdat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim loopc As Long

Dim Miscabezas() As tIndiceCabezaLong


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Head.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If

Dim Leer As New clsIniReader

Call Leer.Initialize(ArchivoAbrir)


Numheads = Val(Leer.GetValue("INIT", "NumHeads"))

ReDim HeadData(0 To Numheads) As HeadData

For i = 1 To Numheads
    InitGrh HeadData(i).Head(1), Val(Leer.GetValue("Head" & i, "Head1")), 0
    InitGrh HeadData(i).Head(2), Val(Leer.GetValue("Head" & i, "Head2")), 0
    InitGrh HeadData(i).Head(3), Val(Leer.GetValue("Head" & i, "Head3")), 0
    InitGrh HeadData(i).Head(4), Val(Leer.GetValue("Head" & i, "Head4")), 0
    DoEvents
    frmMain.LUlitError.Caption = "cabeza: " & i
Next i

Set Leer = Nothing
End Sub


Public Sub CargarEspaldaDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, NumEspalda As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim MisE() As tIndiceCabezaLong


If FileNamePath = vbNullString Then
   If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Capas.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Capas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If

'Resize array
Numcapas = Val(GetVar(ArchivoAbrir, "INIT", "NumCapas"))

ReDim EspaldaAnimData(0 To NumEspalda) As HeadData


For i = 1 To NumEspalda
    InitGrh EspaldaAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa1")), 0
    InitGrh EspaldaAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa2")), 0
    InitGrh EspaldaAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa3")), 0
    InitGrh EspaldaAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Capa" & i, "Capa4")), 0
Next i


End Sub

Public Sub CargarBotasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim Miscabezas() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Botas.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Botas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If

NumBotas = Val(GetVar(ArchivoAbrir, "INIT", "NumBotas"))

'Resize array
ReDim BotasAnimData(0 To NumBotas) As HeadData

For i = 1 To NumBotas
    InitGrh BotasAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Bota" & i, "Bota1")), 0
    InitGrh BotasAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Bota" & i, "Bota2")), 0
    InitGrh BotasAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Bota" & i, "Bota3")), 0
    InitGrh BotasAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Bota" & i, "Bota4")), 0
Next i

End Sub


Public Sub CargarCascosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim ArchivoAbrir As String
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabezaLong


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Cascos.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If


NumCascos = Val(GetVar(ArchivoAbrir, "INIT", "NumCascos"))
'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData

For i = 1 To NumCascos
    InitGrh CascoAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head1")), 0
    InitGrh CascoAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head2")), 0
    InitGrh CascoAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head3")), 0
    InitGrh CascoAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head4")), 0
Next i


End Sub

Public Sub CargarFxsDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer

Dim MisFxs() As tIndiceFxLong
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\fx.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\fx" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If

numfxs = Val(GetVar(ArchivoAbrir, "INIT", "NumFxs"))
ReDim FxData(0 To numfxs) As FxData


For i = 1 To numfxs
    Call InitGrh(FxData(i).Fx, Val(GetVar(ArchivoAbrir, "Fx" & i, "Animacion")), 1)
    FxData(i).OffsetX = Val(GetVar(ArchivoAbrir, "Fx" & i, "OffsetX"))
    FxData(i).OffsetY = Val(GetVar(ArchivoAbrir, "Fx" & i, "OffsetY"))
Next i


End Sub


Public Sub CargarCabezas(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, Index As Integer
Dim ArchivoAbrir As String

Dim Miscabezas() As tIndiceCabeza
Dim MiscabezasLong() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(HeadData()) = 0 Then
        ReDim HeadData(1) As HeadData
    End If
    
    Exit Sub
End If

N = FreeFile
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads


'Resize array
ReDim HeadData(0 To Numheads) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To Numheads
        Get #N, , MiscabezasLong(i)
        InitGrh HeadData(i).Head(1), MiscabezasLong(i).Head(1), 0
        InitGrh HeadData(i).Head(2), MiscabezasLong(i).Head(2), 0
        InitGrh HeadData(i).Head(3), MiscabezasLong(i).Head(3), 0
        InitGrh HeadData(i).Head(4), MiscabezasLong(i).Head(4), 0
    Next i
Else
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i

End If

Close #N

End Sub

Public Sub CargarCascos(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim ArchivoAbrir As String
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza
Dim MiscabezasLong() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(CascoAnimData()) = 0 Then
        ReDim CascoAnimData(1) As HeadData
    End If
    Exit Sub
End If
N = FreeFile
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To NumCascos
        Get #N, , MiscabezasLong(i)
        InitGrh CascoAnimData(i).Head(1), MiscabezasLong(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), MiscabezasLong(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), MiscabezasLong(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), MiscabezasLong(i).Head(4), 0
    Next i
Else
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i
End If

Close #N

End Sub
Public Sub CargarEspalda(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, NumEspalda As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim MisE() As tIndiceCabeza
Dim MisELong() As tIndiceCabezaLong

N = FreeFile

If FileNamePath = vbNullString Then
   If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(EspaldaAnimData()) = 0 Then
        ReDim EspaldaAnimData(1) As HeadData
    End If
    Exit Sub
End If

Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumEspalda

'Resize array
ReDim EspaldaAnimData(0 To NumEspalda) As HeadData
ReDim MisE(0 To NumEspalda + 1) As tIndiceCabeza
ReDim MisELong(0 To NumEspalda + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To NumEspalda
        Get #N, , MisELong(i)
        InitGrh EspaldaAnimData(i).Head(1), MisELong(i).Head(1), 0
        InitGrh EspaldaAnimData(i).Head(2), MisELong(i).Head(2), 0
        InitGrh EspaldaAnimData(i).Head(3), MisELong(i).Head(3), 0
        InitGrh EspaldaAnimData(i).Head(4), MisELong(i).Head(4), 0
    Next i
Else
    For i = 1 To NumEspalda
        Get #N, , MisE(i)
        InitGrh EspaldaAnimData(i).Head(1), MisE(i).Head(1), 0
        InitGrh EspaldaAnimData(i).Head(2), MisE(i).Head(2), 0
        InitGrh EspaldaAnimData(i).Head(3), MisE(i).Head(3), 0
        InitGrh EspaldaAnimData(i).Head(4), MisE(i).Head(4), 0
    Next i
End If
Close #N

End Sub

Public Sub CargarBotas(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim Miscabezas() As tIndiceCabeza
Dim MiscabezasLong() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If


If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(BotasAnimData()) = 0 Then
        ReDim BotasAnimData(1) As HeadData
    End If
    Exit Sub
End If

N = FreeFile
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim BotasAnimData(0 To NumCascos) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To NumCascos + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To NumCascos
        Get #N, , MiscabezasLong(i)
        InitGrh BotasAnimData(i).Head(1), MiscabezasLong(i).Head(1), 0
        InitGrh BotasAnimData(i).Head(2), MiscabezasLong(i).Head(2), 0
        InitGrh BotasAnimData(i).Head(3), MiscabezasLong(i).Head(3), 0
        InitGrh BotasAnimData(i).Head(4), MiscabezasLong(i).Head(4), 0
    Next i
Else
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        InitGrh BotasAnimData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh BotasAnimData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh BotasAnimData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh BotasAnimData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i
End If
Close #N

End Sub


Sub CargarCuerpos(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo
Dim MisCuerposLong() As tIndiceCuerpoLong
Dim ArchivoAbrir As String

N = FreeFile



If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(BodyData()) = 0 Then
        ReDim BodyData(1) As BodyData
    End If
    Exit Sub
End If
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo
ReDim MisCuerposLong(0 To NumCuerpos + 1) As tIndiceCuerpoLong

If UsarGrhLong Then
    For i = 1 To NumCuerpos
        Get #N, , MisCuerposLong(i)
        InitGrh BodyData(i).Walk(1), MisCuerposLong(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerposLong(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerposLong(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerposLong(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerposLong(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerposLong(i).HeadOffsetY
    Next i
Else
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
    Next i
End If

Close #N

End Sub



Public Sub CargarFxs(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer
Dim MisFxs() As tIndiceFx
Dim MisFxslong() As tIndiceFxLong
Dim ArchivoAbrir As String
N = FreeFile


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fx.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    If UBound(FxData()) = 0 Then
        ReDim FxData(1) As FxData
    End If
    Exit Sub
End If
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , numfxs

'Resize array
ReDim FxData(0 To numfxs) As FxData
ReDim MisFxs(0 To numfxs + 1) As tIndiceFx
ReDim MisFxslong(0 To numfxs + 1) As tIndiceFxLong

If UsarGrhLong Then
    For i = 1 To numfxs
        Get #N, , MisFxslong(i)
        Call InitGrh(FxData(i).Fx, MisFxslong(i).Animacion, 1)
        FxData(i).OffsetX = MisFxslong(i).OffsetX
        FxData(i).OffsetY = MisFxslong(i).OffsetY
    Next i
Else
    For i = 1 To numfxs
        Get #N, , MisFxs(i)
        Call InitGrh(FxData(i).Fx, MisFxs(i).Animacion, 1)
        FxData(i).OffsetX = MisFxs(i).OffsetX
        FxData(i).OffsetY = MisFxs(i).OffsetY
    Next i
End If

Close #N

End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************** Funciones de guardado ***********************************
'********************************************************************************
'********************************************************************************

Public Sub GuardarCabezas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cabezas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas

Put #N, , CInt(UBound(HeadData)) 'numheads
Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(HeadData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(HeadData) + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To UBound(HeadData)
        MiscabezasLong(i).Head(1) = HeadData(i).Head(1).grhindex
        MiscabezasLong(i).Head(2) = HeadData(i).Head(2).grhindex
        MiscabezasLong(i).Head(3) = HeadData(i).Head(3).grhindex
        MiscabezasLong(i).Head(4) = HeadData(i).Head(4).grhindex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    For i = 1 To UBound(HeadData)
        Miscabezas(i).Head(1) = HeadData(i).Head(1).grhindex
        Miscabezas(i).Head(2) = HeadData(i).Head(2).grhindex
        Miscabezas(i).Head(3) = HeadData(i).Head(3).grhindex
        Miscabezas(i).Head(4) = HeadData(i).Head(4).grhindex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en cabeza" & i)

End Sub
Public Sub GuardarCabezasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Head.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumHeads", CInt(UBound(HeadData)))

For i = 1 To UBound(HeadData)
    If HeadData(i).Head(1).grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head1", HeadData(i).Head(1).grhindex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head2", HeadData(i).Head(2).grhindex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head3", HeadData(i).Head(3).grhindex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head4", HeadData(i).Head(4).grhindex)
        DoEvents
        frmMain.LUlitError.Caption = "cabeza: " & i
    End If
Next i

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

Exit Sub

ErrHandler:
Call MsgBox("Error en cabeza" & i)

End Sub
Public Sub GuardarFxs(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler
Dim MisFxslong() As tIndiceFxLong

numfxs = UBound(FxData)
ReDim FxDataI(0 To numfxs + 1) As tIndiceFx
ReDim MisFxslong(0 To numfxs + 1) As tIndiceFxLong



Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fx.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Fxs" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N

'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas
Put #N, , numfxs

If UsarGrhLong Then
    For i = 1 To numfxs
        MisFxslong(i).Animacion = FxData(i).Fx.grhindex
        MisFxslong(i).OffsetX = FxData(i).OffsetX
        MisFxslong(i).OffsetY = FxData(i).OffsetY
        Put #N, , MisFxslong(i)
    Next i
Else
    For i = 1 To numfxs
        FxDataI(i).Animacion = FxData(i).Fx.grhindex
        FxDataI(i).OffsetX = FxData(i).OffsetX
        FxDataI(i).OffsetY = FxData(i).OffsetY
        Put #N, , FxDataI(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Fx) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarFxsDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler


numfxs = UBound(FxData)


Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\fx.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\fx" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Call WriteVar(ArchivoAbrir, "INIT", "NumFxs", numfxs)

For i = 1 To numfxs
    If FxData(i).Fx.grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Fx" & i, "Animacion", FxData(i).Fx.grhindex)
        Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetX", FxData(i).OffsetX)
        Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetY", FxData(i).OffsetY)
        frmMain.LUlitError.Caption = "Fx : " & i
        DoEvents
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Fx) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarBotas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Botas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N

Put #N, , MiCabecera

Put #N, , CInt(UBound(BotasAnimData)) 'numheads
Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(BotasAnimData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(BotasAnimData) + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To UBound(BotasAnimData)
        MiscabezasLong(i).Head(1) = BotasAnimData(i).Head(1).grhindex
        MiscabezasLong(i).Head(2) = BotasAnimData(i).Head(2).grhindex
        MiscabezasLong(i).Head(3) = BotasAnimData(i).Head(3).grhindex
        MiscabezasLong(i).Head(4) = BotasAnimData(i).Head(4).grhindex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    For i = 1 To UBound(BotasAnimData)
        Miscabezas(i).Head(1) = BotasAnimData(i).Head(1).grhindex
        Miscabezas(i).Head(2) = BotasAnimData(i).Head(2).grhindex
        Miscabezas(i).Head(3) = BotasAnimData(i).Head(3).grhindex
        Miscabezas(i).Head(4) = BotasAnimData(i).Head(4).grhindex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Botas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en bota " & i)

End Sub

Public Sub GuardarBotasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Botas.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Botas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Call WriteVar(ArchivoAbrir, "INIT", "NumBotas", CInt(UBound(BotasAnimData)))

For i = 1 To UBound(BotasAnimData)
    If BotasAnimData(i).Head(1).grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota1", BotasAnimData(i).Head(1).grhindex)
        Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota2", BotasAnimData(i).Head(2).grhindex)
        Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota3", BotasAnimData(i).Head(3).grhindex)
        Call WriteVar(ArchivoAbrir, "Bota" & i, "Bota4", BotasAnimData(i).Head(4).grhindex)
    End If
Next i

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Botas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en bota " & i)

End Sub

Public Sub GuardarCapas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Capas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N

Put #N, , MiCabecera

Put #N, , CInt(UBound(EspaldaAnimData))  'numheads
Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(EspaldaAnimData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(EspaldaAnimData) + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To UBound(EspaldaAnimData)
        MiscabezasLong(i).Head(1) = EspaldaAnimData(i).Head(1).grhindex
        MiscabezasLong(i).Head(2) = EspaldaAnimData(i).Head(2).grhindex
        MiscabezasLong(i).Head(3) = EspaldaAnimData(i).Head(3).grhindex
        MiscabezasLong(i).Head(4) = EspaldaAnimData(i).Head(4).grhindex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    For i = 1 To UBound(EspaldaAnimData)
        Miscabezas(i).Head(1) = EspaldaAnimData(i).Head(1).grhindex
        Miscabezas(i).Head(2) = EspaldaAnimData(i).Head(2).grhindex
        Miscabezas(i).Head(3) = EspaldaAnimData(i).Head(3).grhindex
        Miscabezas(i).Head(4) = EspaldaAnimData(i).Head(4).grhindex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Capas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en capa " & i)
End Sub

Public Sub GuardarCapasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
   If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Capas.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Capas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If


If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

'Resize array
Call WriteVar(ArchivoAbrir, "INIT", "NumCapas", CInt(UBound(EspaldaAnimData)))

For i = 1 To UBound(EspaldaAnimData)
    If EspaldaAnimData(i).Head(1).grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa1", EspaldaAnimData(i).Head(1).grhindex)
        Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa2", EspaldaAnimData(i).Head(2).grhindex)
        Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa3", EspaldaAnimData(i).Head(3).grhindex)
        Call WriteVar(ArchivoAbrir, "Capa" & i, "Capa4", EspaldaAnimData(i).Head(4).grhindex)
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Capas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en capa " & i)
End Sub

Public Sub GuardarBodys(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Personajes" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas


Dim MisCuerpos() As tIndiceCuerpo
Dim MisCuerposLong() As tIndiceCuerpoLong
ReDim MisCuerpos(0 To UBound(BodyData) + 1) As tIndiceCuerpo

ReDim MisCuerposLong(0 To UBound(BodyData) + 1) As tIndiceCuerpoLong




Put #N, , CInt(UBound(BodyData)) 'numheads

If UsarGrhLong Then
    For i = 1 To UBound(BodyData)
        MisCuerposLong(i).Body(1) = BodyData(i).Walk(1).grhindex
        MisCuerposLong(i).Body(2) = BodyData(i).Walk(2).grhindex
        MisCuerposLong(i).Body(3) = BodyData(i).Walk(3).grhindex
        MisCuerposLong(i).Body(4) = BodyData(i).Walk(4).grhindex
        MisCuerposLong(i).HeadOffsetX = BodyData(i).HeadOffset.X
        MisCuerposLong(i).HeadOffsetY = BodyData(i).HeadOffset.Y
        Put #N, , MisCuerpos(i)
    Next i
Else
    For i = 1 To UBound(BodyData)
        MisCuerpos(i).Body(1) = BodyData(i).Walk(1).grhindex
        MisCuerpos(i).Body(2) = BodyData(i).Walk(2).grhindex
        MisCuerpos(i).Body(3) = BodyData(i).Walk(3).grhindex
        MisCuerpos(i).Body(4) = BodyData(i).Walk(4).grhindex
        MisCuerpos(i).HeadOffsetX = BodyData(i).HeadOffset.X
        MisCuerpos(i).HeadOffsetY = BodyData(i).HeadOffset.Y
        Put #N, , MisCuerpos(i)
    Next i
End If

Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Body) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub
Public Sub GuardarBodysDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Body.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumBodies", CInt(UBound(BodyData))) 'numheads

For i = 1 To UBound(BodyData)
    If BodyData(i).Walk(1).grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK1", BodyData(i).Walk(1).grhindex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK2", BodyData(i).Walk(2).grhindex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK3", BodyData(i).Walk(3).grhindex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK4", BodyData(i).Walk(4).grhindex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsetX", BodyData(i).HeadOffset.X)
        Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsety", BodyData(i).HeadOffset.Y)
        frmMain.LUlitError.Caption = "body : " & i
        DoEvents
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
Exit Sub

EstadoNoGuardado(e_EstadoIndexador.Body) = False

ErrHandler:
Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub
Public Sub GuardarArmas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Narchivo = ArchivoAbrir

Dim esc() As tIndiceCabeza
Dim Nums As Integer
Nums = UBound(WeaponAnimData)
    
ReDim esc(1 To Nums) As tIndiceCabeza

N = FreeFile
Open Narchivo For Binary As #N
    Put #N, , Nums
    
    For i = 1 To UBound(WeaponAnimData)
        esc(i).Head(1) = WeaponAnimData(i).WeaponWalk(1).grhindex
        esc(i).Head(2) = WeaponAnimData(i).WeaponWalk(2).grhindex
        esc(i).Head(3) = WeaponAnimData(i).WeaponWalk(3).grhindex
        esc(i).Head(4) = WeaponAnimData(i).WeaponWalk(4).grhindex
        
        Put #N, , esc(i)
    Next i
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Armas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en arma " & i)
End Sub
Public Sub GuardarEscudos(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Dim esc() As tIndiceCabeza
Dim Nums As Integer
Nums = UBound(ShieldAnimData)
    
ReDim esc(1 To Nums) As tIndiceCabeza

N = FreeFile
Open ArchivoAbrir For Binary As #N
    Put #N, , Nums
    
    For i = 1 To UBound(ShieldAnimData)
        esc(i).Head(1) = ShieldAnimData(i).ShieldWalk(1).grhindex
        esc(i).Head(2) = ShieldAnimData(i).ShieldWalk(2).grhindex
        esc(i).Head(3) = ShieldAnimData(i).ShieldWalk(3).grhindex
        esc(i).Head(4) = ShieldAnimData(i).ShieldWalk(4).grhindex
        
        Put #N, , esc(i)
    Next i
Close #N
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Escudos) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en escudo " & i)
End Sub
Public Sub GuardarCascos(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos.ind"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Cascos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas

Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(CascoAnimData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(CascoAnimData) + 1) As tIndiceCabezaLong

Put #N, , CInt(UBound(CascoAnimData)) 'numheads

If UsarGrhLong Then
    For i = 1 To UBound(CascoAnimData)
        MiscabezasLong(i).Head(1) = CascoAnimData(i).Head(1).grhindex
        MiscabezasLong(i).Head(2) = CascoAnimData(i).Head(2).grhindex
        MiscabezasLong(i).Head(3) = CascoAnimData(i).Head(3).grhindex
        MiscabezasLong(i).Head(4) = CascoAnimData(i).Head(4).grhindex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    
    For i = 1 To UBound(CascoAnimData)
        Miscabezas(i).Head(1) = CascoAnimData(i).Head(1).grhindex
        Miscabezas(i).Head(2) = CascoAnimData(i).Head(2).grhindex
        Miscabezas(i).Head(3) = CascoAnimData(i).Head(3).grhindex
        Miscabezas(i).Head(4) = CascoAnimData(i).Head(4).grhindex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en casco " & i)

End Sub

Public Sub GuardarCascosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Cascos.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumCascos", CInt(UBound(CascoAnimData)))

For i = 1 To UBound(CascoAnimData)
    If CascoAnimData(i).Head(1).grhindex > 0 Then
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head1", CascoAnimData(i).Head(1).grhindex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head2", CascoAnimData(i).Head(2).grhindex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head3", CascoAnimData(i).Head(3).grhindex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head4", CascoAnimData(i).Head(4).grhindex)
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en casco " & i)

End Sub


Public Sub GuardarArmasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas.dat"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\armas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Narchivo = ArchivoAbrir
Call WriteVar(Narchivo, "INIT", "NumArmas", UBound(WeaponAnimData))
For i = 1 To UBound(WeaponAnimData)
    If WeaponAnimData(i).WeaponWalk(1).grhindex > 0 Then
        Call WriteVar(Narchivo, "ARMA" & i, "Dir1", WeaponAnimData(i).WeaponWalk(1).grhindex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir2", WeaponAnimData(i).WeaponWalk(2).grhindex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir3", WeaponAnimData(i).WeaponWalk(3).grhindex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir4", WeaponAnimData(i).WeaponWalk(4).grhindex)
    End If
Next i
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Armas) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en arma " & i)
End Sub
Public Sub GuardarEscudosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo ErrHandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos.dat"
    Else
        ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\Escudos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Narchivo = ArchivoAbrir
Call WriteVar(Narchivo, "INIT", "NumEscudos", UBound(ShieldAnimData))
For i = 1 To UBound(ShieldAnimData)
    If ShieldAnimData(i).ShieldWalk(1).grhindex > 0 Then
        Call WriteVar(Narchivo, "ESC" & i, "Dir1", ShieldAnimData(i).ShieldWalk(1).grhindex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir2", ShieldAnimData(i).ShieldWalk(2).grhindex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir3", ShieldAnimData(i).ShieldWalk(3).grhindex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir4", ShieldAnimData(i).ShieldWalk(4).grhindex)
    End If
Next i
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Escudos) = False

Exit Sub
ErrHandler:
Call MsgBox("Error en escudo " & i)
End Sub


'********************************************************************************
'********************************************************************************
'********************************************************************************
'******************************** Botones ***************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************

Public Sub BotonGuardado(Optional ByVal FileNamePath As String = vbNullString)

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call SaveGrhData(FileNamePath)
    Case e_EstadoIndexador.Body
        Call GuardarBodys(FileNamePath)
    Case e_EstadoIndexador.Cabezas
        Call GuardarCabezas(FileNamePath)
    Case e_EstadoIndexador.Cascos
        Call GuardarCascos(FileNamePath)
    Case e_EstadoIndexador.Escudos
        Call GuardarEscudos(FileNamePath)
    Case e_EstadoIndexador.Armas
        Call GuardarArmas(FileNamePath)
    Case e_EstadoIndexador.Botas
        Call GuardarBotas(FileNamePath)
    Case e_EstadoIndexador.Capas
        Call GuardarCapas(FileNamePath)
    Case e_EstadoIndexador.Fx
        Call GuardarFxs(FileNamePath)
End Select
End Sub
Public Sub BotonGuardadoDat(Optional ByVal FileNamePath As String = vbNullString)

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call SaveGrhDataDat(FileNamePath)
    Case e_EstadoIndexador.Body
        Call GuardarBodysDat(FileNamePath)
    Case e_EstadoIndexador.Cabezas
        Call GuardarCabezasDat(FileNamePath)
    Case e_EstadoIndexador.Cascos
        Call GuardarCascosDat(FileNamePath)
    Case e_EstadoIndexador.Escudos
        Call GuardarEscudosDat(FileNamePath)
    Case e_EstadoIndexador.Armas
        Call GuardarArmasDat(FileNamePath)
    Case e_EstadoIndexador.Botas
        Call GuardarBotasDat(FileNamePath)
    Case e_EstadoIndexador.Capas
        Call GuardarCapasDat(FileNamePath)
    Case e_EstadoIndexador.Fx
        Call GuardarFxsDat(FileNamePath)
End Select
End Sub
Public Sub BotonCargado(Optional ByVal FileNamePath As String = vbNullString)
Dim respuesta As Byte
Dim tempLong As Long

respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")
If respuesta <> vbYes Then
    Exit Sub
End If
        
frmMain.Visor.Cls
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call LoadGrhData(FileNamePath)
        Call RenuevaListaGrH
    Case e_EstadoIndexador.Body
        Call CargarCuerpos(FileNamePath)
        Call RenuevaListaBodys
    Case e_EstadoIndexador.Cabezas
        Call CargarCabezas(FileNamePath)
        Call RenuevaListaCabezas
    Case e_EstadoIndexador.Cascos
        Call CargarCascos(FileNamePath)
        Call RenuevaListaCascos
    Case e_EstadoIndexador.Escudos
        Call CargarAnimEscudos(FileNamePath)
        Call RenuevaListaEscudos
    Case e_EstadoIndexador.Armas
        Call CargarAnimArmas(FileNamePath)
        Call RenuevaListaArmas
    Case e_EstadoIndexador.Botas
        Call CargarBotas(FileNamePath)
        Call RenuevaListaBotas
    Case e_EstadoIndexador.Capas
        Call CargarEspalda(FileNamePath)
        Call RenuevaListaCapas
    Case e_EstadoIndexador.Fx
        Call CargarFxs(FileNamePath)
        Call RenuevaListaFX
End Select
If EstadoIndexador = e_EstadoIndexador.Grh Then
    tempLong = ListaindexGrH(GRHActual)
    If tempLong >= frmMain.Lista.ListCount Then tempLong = 0
    frmMain.Lista.listIndex = tempLong
Else
    tempLong = ListaindexGrH(DataIndexActual)
    If tempLong >= frmMain.Lista.ListCount Then tempLong = 0
    frmMain.Lista.listIndex = tempLong
End If
End Sub

Public Sub BotonCargadoDat(Optional ByVal FileNamePath As String = vbNullString)
Dim respuesta As Byte
Dim tempLong As Long

respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")
If respuesta <> vbYes Then
    Exit Sub
End If
        
frmMain.Visor.Cls
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call LoadGrhDataDat(FileNamePath)
        Call RenuevaListaGrH
    Case e_EstadoIndexador.Body
        Call CargarCuerposDat(FileNamePath)
        Call RenuevaListaBodys
    Case e_EstadoIndexador.Cabezas
        Call CargarCabezasdat(FileNamePath)
        Call RenuevaListaCabezas
    Case e_EstadoIndexador.Cascos
        Call CargarCascosDat(FileNamePath)
        Call RenuevaListaCascos
    Case e_EstadoIndexador.Escudos
        Call CargarAnimEscudos(FileNamePath)
        Call RenuevaListaEscudos
    Case e_EstadoIndexador.Armas
        Call CargarAnimArmas(FileNamePath)
        Call RenuevaListaArmas
    Case e_EstadoIndexador.Botas
        Call CargarBotasDat(FileNamePath)
        Call RenuevaListaBotas
    Case e_EstadoIndexador.Capas
        Call CargarEspaldaDat(FileNamePath)
        Call RenuevaListaCapas
    Case e_EstadoIndexador.Fx
        Call CargarFxsDat(FileNamePath)
        Call RenuevaListaFX
End Select
If EstadoIndexador = e_EstadoIndexador.Grh Then
    tempLong = ListaindexGrH(GRHActual)
    frmMain.Lista.listIndex = tempLong
Else
    tempLong = ListaindexGrH(DataIndexActual)
    frmMain.Lista.listIndex = tempLong
End If
End Sub

