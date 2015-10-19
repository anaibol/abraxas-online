Attribute VB_Name = "modGeneral"
Option Explicit

Public AmbientColor As D3DCOLORVALUE

Public Type WorldPos
    map As Integer
    x As Integer
    y As Integer
End Type
    

Private Type BCOLOR
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type
    
Public Type ParticleSave
    equation As Byte
    name As String * 32
    ratio As Integer
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
        
    textures(1 To 10) As Integer
    numTexs As Integer
    
    Life As Integer
    
    nParticles As Long
    
    size As Integer
    Gravity As Single
    
    vX As Single
    vY As Single

    alpha As Byte
    
    sColor As BCOLOR
    eColor As BCOLOR
End Type

Public NumParticles As Integer
Public pSaves() As ParticleSave

Public UltimoClickX As Integer
Public UltimoClickY As Integer
Type SupData
    name As String
    Grh As Integer
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte
End Type
Public MaxSup As Integer
Public SupData() As SupData

Public ShadowRGB(0 To 3) As Long
Public AlphaRGB(0 To 3) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Function OpenFolderSearch(Optional Caption As String = "") As String
Dim SheN As New Shell32.Shell, Fo As Folder2, FDir As FolderItem
    On Local Error Resume Next
    If Caption = "" Then Caption = "Abraxas Map Editor" 'App.EXEName
    Set Fo = SheN.BrowseForFolder(frmMain.hwnd, Caption, 0&)
    Set FDir = Fo.Self
    OpenFolderSearch = FDir.Path
    Set Fo = Nothing
    Set FDir = Nothing
End Function
Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < 1 Or x > MapInfo.dX Or y < 1 Or y > MapInfo.dY Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
Sub Main()
On Error GoTo err
    ChDir App.Path
    ChDrive App.Path
       
    If Len(DirDat) < 2 Or Len(DirGraficos) < 2 Then
        frmConfig.Show , frmMain
        Exit Sub
    End If
    
    Call Audio.Initialize(frmMain.hwnd, DirSound, DirMidi)

    AmbientColor.a = 255
    AmbientColor.r = 255
    AmbientColor.g = 255
    AmbientColor.b = 255

    Map_New
    
    frmLoad.Show , frmMain
    frmLoad.Shape1.Width = 1
    
        frmLoad.Label1.Caption = "Iniciando Engine"
        Call Engine.Engine_Init
        base_light = General_RGB_Color_to_Long(200, 200, 200, 255)
        TmpRGB(0) = -1
        TmpRGB(1) = -1
        TmpRGB(2) = -1
        TmpRGB(3) = -1
        
        ShadowRGB(0) = 1677721600
        ShadowRGB(1) = 1677721600
        ShadowRGB(2) = 1677721600
        ShadowRGB(3) = 1677721600
        
        GrillRGB(0) = D3DColorARGB(100, 255, 0, 255)
        GrillRGB(1) = D3DColorARGB(100, 255, 0, 255)
        GrillRGB(2) = D3DColorARGB(100, 255, 0, 255)
        GrillRGB(3) = D3DColorARGB(100, 255, 0, 255)
        
        AlphaRGB(0) = D3DColorARGB(100, 255, 255, 255)
        AlphaRGB(1) = D3DColorARGB(100, 255, 255, 255)
        AlphaRGB(2) = D3DColorARGB(100, 255, 255, 255)
        AlphaRGB(3) = D3DColorARGB(100, 255, 255, 255)
        
        ShowLayer1 = True
        ShowLayer2 = True
        ShowLayer3 = True
        ShowObjs = True
        ShowNpcs = True
        ShowTrans = True

        cData.cBloq = True
        cData.cTrig = True
        cData.cObj = True
        cData.cNpc = True

        Dim i As Long
        For i = 1 To 4
            cData.cCap(i) = True
        Next i
        
        Load frmMode
        
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 80
    
    frmLoad.Label1.Caption = "Cargando GRHs"
        modLoadData.LoadGrhData
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 20
    
    frmLoad.Label1.Caption = "Cargando Cabezas"
        modLoadData.CargarCabezas
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 20
    frmLoad.Label1.Caption = "Cargando IndicesSuperficie"
        modLoadData.CargarIndicesSuperficie
    frmLoad.Label1.Caption = "Cargando Cuerpos"
        modLoadData.CargarCuerpos
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 20
 
    frmLoad.Label1.Caption = "Cargando Cascos"
        modLoadData.CargarCascos
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 20

    frmLoad.Label1.Caption = "Cargando FXs"
        modLoadData.CargarFxs
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 50
    
    frmLoad.Label1.Caption = "Cargando Objetos"
    LoadOBJData
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 160
    
    frmLoad.Label1.Caption = "Cargando NPCs"
    CargarIndicesNPC
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 80

    
    frmLoad.Label1.Caption = "Carga Terminada, Bienvenido."
    frmLoad.Shape1.Width = frmLoad.Shape1.Width + 30
    
        Sleep 500
        Unload frmLoad
    
    Load frmMain
    frmMain.Show
    frmMode.Show , frmMain
    frmMinimap.Show , frmMain
    
    frmMode.optRes_Click 2
    frmMode.optVel_Click 2
    
    DoEvents
    
    Deshacer_Clear
    
    Do While bRunning
        If frmMain.WindowState <> vbMinimized Then
            CheckKeys
            Engine.Render
            DoEvents
        Else
            DoEvents
            Sleep 100&
        End If
    Loop

    Engine.Engine_Deinit
    Exit Sub
err:
    Open App.Path & "\errores.log" For Binary As #1
        Put #1, , CStr(err.Number & " " & err.Description) & vbCrLf
    Close #1
    
    Resume Next
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    If MiMouse = False Then Exit Sub
    
'If UserMoving = 1 Then Exit Sub
    If UserMoving = 0 Then
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            Engine.Engine_MoveScreen NORTH
            Exit Sub
        End If

        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            Engine.Engine_MoveScreen EAST
            Exit Sub
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            Engine.Engine_MoveScreen SOUTH
            Exit Sub
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            Engine.Engine_MoveScreen WEST
            Exit Sub
        End If
    End If
End Sub
Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
Function DirCliente() As String
    DirCliente = General_Var_Get(App.Path & "\MapEditor.ini", "CONFIG", "Cliente")
    If Right$(DirCliente, 1) <> "\" Then
        DirCliente = DirCliente & "\"
    End If
End Function
Function DirGraficos() As String
    DirGraficos = DirCliente & "Grh\"
End Function
Function DirSound() As String
    DirSound = DirCliente & "Wavs\"
End Function
Function DirMidi() As String
    DirMidi = DirCliente & "Midi\"
End Function
Function DirInit() As String
    DirInit = DirCliente & "Data\"
End Function
Function DirDat() As String
    DirDat = General_Var_Get(App.Path & "\MapEditor.ini", "CONFIG", "Dats")
    If Right$(DirDat, 1) <> "\" Then
        DirDat = DirDat & "\"
    End If
End Function

Public Function ReadFieldLen(Text As String, SepASCII As Integer) As Long
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            LastPos = i
        End If
    Next i
    
    ReadFieldLen = LastPos
End Function

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = mid(Text, LastPos + 1)
End If

End Function

Sub DameArchivo(ByVal dstr As String, ByRef file As String, ByRef direc As String)
    Dim miLen As Long
    miLen = ReadFieldLen(dstr, Asc("\"))
    
    file = mid$(dstr, miLen + 1, Len(dstr) - miLen)
    direc = mid$(dstr, 1, miLen)
End Sub
