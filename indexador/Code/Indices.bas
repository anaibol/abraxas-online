Attribute VB_Name = "Indices"
Option Explicit


Public CascoSData() As tIndiceCabeza
Public CapasData() As tIndiceCabeza
Public BotasData() As tIndiceCabeza
Public headataI() As tIndiceCabeza
Public Mapas() As Byte
Public CuerpoData() As tIndiceCuerpo
Public FxDataI() As tIndiceFx



Public Numheads As Integer
Public NumCascos As Integer
Public NumBotas As Integer
Public Numcapas As Integer
Public NumCuerpos As Integer
Public NumTips As Integer
Public NumMapas As Integer

Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Function GetVar(file As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
Sub CargarAnimArmas(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next

    Dim loopc As Long
    Dim ArchivoAbrir As String
    Dim N As Long
    
    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas" & SavePath & ".ind"
        End If
    Else
        ArchivoAbrir = FileNamePath
    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        If UBound(WeaponAnimData()) = 0 Then
            ReDim WeaponAnimData(1) As WeaponAnimData
        End If
        Exit Sub
    End If
    
    Dim esc() As tIndiceCabeza
    
    N = FreeFile
    Open ArchivoAbrir For Binary As #N
        Get #N, , NumWeaponAnims
        
        ReDim esc(1 To NumWeaponAnims) As tIndiceCabeza
        ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
        
        For loopc = 1 To NumWeaponAnims
            Get #N, , esc(loopc)
            
            WeaponAnimData(loopc).WeaponWalk(1).grhindex = esc(loopc).Head(1)
            WeaponAnimData(loopc).WeaponWalk(2).grhindex = esc(loopc).Head(2)
            WeaponAnimData(loopc).WeaponWalk(3).grhindex = esc(loopc).Head(3)
            WeaponAnimData(loopc).WeaponWalk(4).grhindex = esc(loopc).Head(4)
        Next loopc
    Close #N
End Sub
Sub CargarAnimEscudos(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
    Dim loopc As Long
    Dim ArchivoAbrir As String
    Dim N As Long

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos.ind"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos" & SavePath & ".ind"
        End If
    Else
        ArchivoAbrir = FileNamePath
    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        If UBound(ShieldAnimData()) = 0 Then
            ReDim ShieldAnimData(1) As ShieldAnimData
        End If
        Exit Sub
    End If

    Dim esc() As tIndiceCabeza
    
    N = FreeFile
    Open ArchivoAbrir For Binary As #N
        Get #N, , NumEscudosAnims
        
        ReDim esc(1 To NumEscudosAnims) As tIndiceCabeza
        ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
        
        For loopc = 1 To NumEscudosAnims
            Get #N, , esc(loopc)
            
            ShieldAnimData(loopc).ShieldWalk(1).grhindex = esc(loopc).Head(1)
            ShieldAnimData(loopc).ShieldWalk(2).grhindex = esc(loopc).Head(2)
            ShieldAnimData(loopc).ShieldWalk(3).grhindex = esc(loopc).Head(3)
            ShieldAnimData(loopc).ShieldWalk(4).grhindex = esc(loopc).Head(4)
        Next loopc
    Close #N
End Sub
Sub CargarAnimArmasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    Dim ArchivoAbrir As String
    
    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "armas" & SavePath & ".dat"
        End If
    Else
        ArchivoAbrir = FileNamePath
    End If

    
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        If UBound(WeaponAnimData()) = 0 Then
            ReDim WeaponAnimData(1) As WeaponAnimData
        End If
        Exit Sub
    End If
    
    arch = ArchivoAbrir
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub
Sub CargarAnimEscudosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    Dim ArchivoAbrir As String
    

    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos.dat"
        Else
            ArchivoAbrir = App.Path & "\" & CarpetaDeInis & "\" & "escudos" & SavePath & ".dat"
        End If
    Else
        ArchivoAbrir = FileNamePath
    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        If UBound(ShieldAnimData()) = 0 Then
            ReDim ShieldAnimData(1) As ShieldAnimData
        End If
        Exit Sub
    End If
    
    arch = ArchivoAbrir
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Sub LoadGrhDataDat(Optional ByVal FileNamePath As String = vbNullString)
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer
Dim ArchivoAbrir As String
Dim StringGrh As String


'Resize arrays
ReDim GrhData(1 To MAXGrH) As GrhData

'Open files

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = IniPath & "Graficos.dat"
    Else
        ArchivoAbrir = IniPath & "Graficos" & SavePath & ".dat"
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

Do Until Grh > MAXGrH
    
    'Get number of frames
    StringGrh = Leer.GetValue("Graphics", "Grh" & Grh)
    If StringGrh <> vbNullString Then
        
        GrhData(Grh).NumFrames = Val(ReadField(1, StringGrh, Asc("-")))
    
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        If GrhData(Grh).NumFrames > 1 Then
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                GrhData(Grh).Frames(Frame) = Val(ReadField(1 + Frame, StringGrh, Asc("-")))
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > MAXGrH Then
                    GoTo ErrorHandler
                End If
            
            Next Frame
        
            GrhData(Grh).speed = Val(ReadField(1 + Frame, StringGrh, Asc("-")))

            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        
        Else
            'Read in normal GRH data
            GrhData(Grh).FileNum = Val(ReadField(2, StringGrh, Asc("-")))
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
    
             GrhData(Grh).sX = Val(ReadField(3, StringGrh, Asc("-")))

            
             GrhData(Grh).sY = Val(ReadField(4, StringGrh, Asc("-")))

                
            GrhData(Grh).pixelWidth = Val(ReadField(5, StringGrh, Asc("-")))

            GrhData(Grh).pixelHeight = Val(ReadField(6, StringGrh, Asc("-")))
            
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
            
            GrhData(Grh).Frames(1) = Grh
                
        End If
        
    End If
    'Get Next Grh Number
    Grh = Grh + 1

Loop
'************************************************
Set Leer = Nothing

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading " & ArchivoAbrir & " Stopped at GRH number: " & Grh

End Sub

Sub SaveGrhData(Optional ByVal FileNamePath As String = vbNullString)
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Long
Dim Frame As Integer
Dim TempInt As Integer
Dim N As Integer
Dim ArchivoAbrir As String
Dim fileVersion As Long

fileVersion = 1539209

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = IniPath & "Grh.ind"
    Else
        ArchivoAbrir = IniPath & "Grh" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

N = FreeFile

'Open files
Open ArchivoAbrir For Binary As #N
Seek #1, 1

Dim grhCount As Long
grhCount = UBound(GrhData)

Put #1, , fileVersion
Put #1, , grhCount

'Fill Grh List
For Grh = 1 To MAXGrH
    If GrhData(Grh).NumFrames <= 0 Then GoTo aqui2
    
    'Get first Grh Number
    Put #1, , Grh

    With GrhData(Grh)
        'Get number of frames
        Put #1, , GrhData(Grh).NumFrames
    
        If .NumFrames > 1 Then
            'Read a animation GRH set
            For Frame = 1 To .NumFrames
                Put #1, , .Frames(Frame)
            Next Frame
            
            If .speed <= 0 Then .speed = 1
            Put #1, , .speed
            
        Else
            'Read in normal GRH data
            Put #1, , .FileNum
            
            Put #1, , .sX
            Put #1, , .sY

            Put #1, , .pixelWidth
            Put #1, , .pixelHeight
        End If
    End With
    'Get Next Grh Number
aqui2:
Next Grh
'************************************************

Close #1

 EstadoNoGuardado(e_EstadoIndexador.Grh) = False
 frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while saving the " & ArchivoAbrir & " ! Stopped at GRH number: " & Grh

End Sub

Sub SaveGrhDataDat(Optional ByVal FileNamePath As String = vbNullString)
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim TempInt As Integer
Dim ArchivoAbrir As String
Dim StringGrh As String
Dim LastGrh As Long
Dim TotalString As String

'Resize arrays

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = IniPath & "Graficos.dat"
    Else
        ArchivoAbrir = IniPath & "Graficos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


'TotalString = "[Graphics]" & vbCrLf & vbCrLf
Grh = 1
Do Until Grh > MAXGrH
    
    'Get number of frames
    If GrhData(Grh).NumFrames >= 1 Then
        StringGrh = GrhData(Grh).NumFrames & "-"

        If GrhData(Grh).NumFrames > 1 Then
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                StringGrh = StringGrh & GrhData(Grh).Frames(Frame) & "-"
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > MAXGrH Then
                    GoTo ErrorHandler
                End If
            
            Next Frame
        
            StringGrh = StringGrh & GrhData(Grh).speed
            If GrhData(Grh).speed <= 0 Then GoTo ErrorHandler
        Else
            'Read in normal GRH data
            StringGrh = StringGrh & GrhData(Grh).FileNum & "-"
            
    
            StringGrh = StringGrh & GrhData(Grh).sX & "-"
        
            StringGrh = StringGrh & GrhData(Grh).sY & "-"
                
            StringGrh = StringGrh & GrhData(Grh).pixelWidth & "-"
            
            StringGrh = StringGrh & GrhData(Grh).pixelHeight
        End If
        Call WriteVar(ArchivoAbrir, "Graphics", "Grh" & Grh, StringGrh)
        'TotalString = TotalString & "Grh" & Grh & "=" & StringGrh & vbCrLf
        LastGrh = Grh
        DoEvents
    End If
    'Get Next Grh Number
    Grh = Grh + 1

    frmMain.LUlitError.Caption = "Grh: " & Grh
Loop
'************************************************
Call WriteVar(ArchivoAbrir, "INIT", "NumGrh", LastGrh)
'TotalString = TotalString & vbCrLf & "[INIT]" & vbCrLf & "numGRH" & "=" & LastGrh

'    Dim N As Integer
'    N = FreeFile
    
'    Open ArchivoAbrir For Binary As #N
'        Put #N, , TotalString
'    Close #N

EstadoNoGuardado(e_EstadoIndexador.Grh) = False
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
Exit Sub




ErrorHandler:

MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Public Function ListaindexGrH(ByVal numGRH As Integer) As Integer
Dim i As Long
ListaindexGrH = -1
For i = 0 To frmMain.Lista.ListCount
    If numGRH = Val(ReadField(1, frmMain.Lista.List(i), Asc(" "))) Then
        ListaindexGrH = i
        Exit Function
    End If
Next i

End Function

Public Function ComprobarSobreescribir(ByVal ArchivoAbrir As String) As Boolean
' Comprueba si el archvo existe y advierte de sobreescritura. Si se acepta ya lo borra

    If FileExist(ArchivoAbrir, vbArchive) Then
        Dim respuesta As Byte
        respuesta = MsgBox("ATENCION Si contunias sobrescribiras el archivo existente" & vbCrLf & ArchivoAbrir, 4, "¡¡ADVERTENCIA!!")
        If respuesta <> vbYes Then
            ComprobarSobreescribir = False
            Exit Function
        End If
        Kill ArchivoAbrir
    End If
    ComprobarSobreescribir = True
End Function


Public Sub ComprobarIndexLista()

    If UltimoindexE(EstadoIndexador) < 0 Then
        If UltimoindexE(EstadoIndexador) <> -1 Then
            frmMain.Lista.listIndex = 0
        Else
            frmMain.Lista.listIndex = -1
        End If
    ElseIf UltimoindexE(EstadoIndexador) >= frmMain.Lista.ListCount Then
        frmMain.Lista.listIndex = frmMain.Lista.ListCount - 1
    Else
        frmMain.Lista.listIndex = UltimoindexE(EstadoIndexador)
    End If

End Sub


Public Function BuscarGrHlibre() As Integer
Dim i As Long
For i = 1 To MAXGrH
    If GrhData(i).NumFrames = 0 Then
        BuscarGrHlibre = i
        Exit Function
    End If
Next i
End Function


Public Function BuscarGrHlibres(ByVal hTotales As Integer) As Integer
Dim i As Long
Dim Primero As Integer
Dim Cuenta As Integer

For i = 1 To MAXGrH
    If Cuenta = hTotales Then
        BuscarGrHlibres = Primero
        Exit Function
    End If
    If GrhData(i).NumFrames = 0 Then
        If Primero = 0 Then
            Primero = i
            Cuenta = 1
        Else
            Cuenta = Cuenta + 1
        End If
    Else
        Cuenta = 0
        Primero = 0
    End If
Next i

End Function


Public Function hayGrHlibres(ByVal Primero As Integer, ByVal hTotales As Integer) As Boolean
Dim i As Long
Dim Cuenta As Integer
If Primero <= 0 Or Primero > MAXGrH Then Exit Function

For i = Primero To Primero + hTotales - 1
    If GrhData(i).NumFrames > 0 Then
        hayGrHlibres = False
        Exit Function
    End If
Next i
hayGrHlibres = True
End Function
Public Sub AgregaGrHex(ByVal numGRH As Integer)

Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

GrhData(numGRH).FileNum = 1
GrhData(numGRH).NumFrames = 1
GrhData(numGRH).pixelHeight = 32
GrhData(numGRH).pixelWidth = 32
ReDim GrhData(numGRH).Frames(1 To 1) As Long
GrhData(numGRH).Frames(1) = numGRH
End Sub
Public Sub AgregaGrH(ByVal numGRH As Integer)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

GrhData(numGRH).FileNum = 1
GrhData(numGRH).NumFrames = 1
GrhData(numGRH).pixelHeight = 32
GrhData(numGRH).pixelWidth = 32
ReDim GrhData(numGRH).Frames(1 To 1) As Long
GrhData(numGRH).Frames(1) = numGRH

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To MAXGrH
    If GrhData(i).NumFrames = 1 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    ElseIf GrhData(i).NumFrames > 1 Then
        frmMain.Lista.AddItem i & " (animacion)"
        CuentaIndex = CuentaIndex + 1
    End If
    If i = numGRH Then
        EsteIndex = CuentaIndex
    End If
Next i
frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaBody(ByVal Numbody As Integer, Optional ByVal RefreshList As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If Numbody > UBound(BodyData) Then ReDim Preserve BodyData(0 To Numbody) As BodyData

BodyData(Numbody).Walk(1).grhindex = 1

If Not RefreshList Then Exit Sub

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(BodyData)
    If BodyData(i).Walk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = Numbody Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub mueveBody(ByVal Numbody As Integer, ByVal origenBody As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim BodyVacio As BodyData
Dim respuesta As Byte

If Numbody > UBound(BodyData) Then ReDim Preserve BodyData(0 To Numbody) As BodyData
If BodyData(Numbody).Walk(1).grhindex > 0 Then
    respuesta = MsgBox("El body " & Numbody & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")
    If respuesta = vbYes Then
        BodyData(Numbody) = BodyData(origenBody)
        If BorrarOriginal Then BodyData(origenBody) = BodyVacio
    End If
Else
    BodyData(Numbody) = BodyData(origenBody)
    If BorrarOriginal Then BodyData(origenBody) = BodyVacio
End If

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(BodyData)
    If BodyData(i).Walk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = Numbody Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveCabeza(ByVal NumHead As Integer, ByVal origenHead As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim headVacia As HeadData

If NumHead > UBound(HeadData) Then ReDim Preserve HeadData(0 To NumHead) As HeadData
If HeadData(NumHead).Head(1).grhindex > 0 Then
    respuesta = MsgBox("La cabeza " & NumHead & " ya existe, ¿deseas sobreescribirla?", 4, "Aviso")
    If respuesta = vbYes Then
        HeadData(NumHead) = HeadData(origenHead)
        If BorrarOriginal Then HeadData(origenHead) = headVacia
    End If
Else
    HeadData(NumHead) = HeadData(origenHead)
    If BorrarOriginal Then HeadData(origenHead) = headVacia
End If


CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(HeadData)
    If HeadData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumHead Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaCabeza(ByVal NumHead As Integer)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumHead > UBound(HeadData) Then ReDim Preserve HeadData(0 To NumHead) As HeadData

HeadData(NumHead).Head(1).grhindex = 1

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(HeadData)
    If HeadData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumHead Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaCasco(ByVal NumCasco As Integer)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumCasco > UBound(CascoAnimData) Then ReDim Preserve CascoAnimData(0 To NumCasco) As HeadData

CascoAnimData(NumCasco).Head(1).grhindex = 1

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(CascoAnimData)
    If CascoAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumCasco Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveCasco(ByVal NumCasco As Integer, ByVal OrigenCasco As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim headVacia As HeadData

If NumCasco > UBound(CascoAnimData) Then ReDim Preserve CascoAnimData(0 To NumCasco) As HeadData

If CascoAnimData(NumCasco).Head(1).grhindex > 0 Then
    respuesta = MsgBox("El casco " & NumCasco & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")
    If respuesta = vbYes Then
        CascoAnimData(NumCasco) = CascoAnimData(OrigenCasco)
        If BorrarOriginal Then CascoAnimData(OrigenCasco) = headVacia
    End If
Else
    CascoAnimData(NumCasco) = CascoAnimData(OrigenCasco)
    If BorrarOriginal Then CascoAnimData(OrigenCasco) = headVacia
End If

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(CascoAnimData)
    If CascoAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumCasco Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveEscudo(ByVal NumEscudo As Integer, ByVal origenEscudo As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim escudoVacio As ShieldAnimData
escudoVacio.ShieldWalk(1).grhindex = 0

If NumEscudo > UBound(ShieldAnimData) Then ReDim Preserve ShieldAnimData(1 To NumEscudo) As ShieldAnimData


If ShieldAnimData(NumEscudo).ShieldWalk(1).grhindex > 0 Then
    respuesta = MsgBox("El escudo " & NumEscudo & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")
    If respuesta = vbYes Then
        ShieldAnimData(NumEscudo) = ShieldAnimData(origenEscudo)
        If BorrarOriginal Then ShieldAnimData(origenEscudo) = escudoVacio
    End If
Else
    ShieldAnimData(NumEscudo) = ShieldAnimData(origenEscudo)
    If BorrarOriginal Then ShieldAnimData(origenEscudo) = escudoVacio
End If

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(ShieldAnimData)
    If ShieldAnimData(i).ShieldWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumEscudo Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaEscudo(ByVal NumEscudo As Integer, Optional ByVal RefreshList As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumEscudo > UBound(ShieldAnimData) Then ReDim Preserve ShieldAnimData(1 To NumEscudo) As ShieldAnimData

ShieldAnimData(NumEscudo).ShieldWalk(1).grhindex = 1

If Not RefreshList Then Exit Sub
CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(ShieldAnimData)
    If ShieldAnimData(i).ShieldWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumEscudo Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaArma(ByVal NumArma As Integer, Optional ByVal RefreshList As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumArma > UBound(WeaponAnimData) Then ReDim Preserve WeaponAnimData(1 To NumArma) As WeaponAnimData

WeaponAnimData(NumArma).WeaponWalk(1).grhindex = 1

If Not RefreshList Then Exit Sub
CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(WeaponAnimData)
    If WeaponAnimData(i).WeaponWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumArma Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveArma(ByVal NumArma As Integer, ByVal OrigenArma As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim armaVacia As WeaponAnimData
armaVacia.WeaponWalk(1).grhindex = 0
If NumArma > UBound(WeaponAnimData) Then ReDim Preserve WeaponAnimData(1 To NumArma) As WeaponAnimData

If WeaponAnimData(NumArma).WeaponWalk(1).grhindex > 0 Then
    respuesta = MsgBox("El arma " & NumArma & " ya existe, ¿deseas sobreescribirla?", 4, "Aviso")
    If respuesta = vbYes Then
        WeaponAnimData(NumArma) = WeaponAnimData(OrigenArma)
        If BorrarOriginal Then WeaponAnimData(OrigenArma) = armaVacia
    End If
Else
    WeaponAnimData(NumArma) = WeaponAnimData(OrigenArma)
    If BorrarOriginal Then WeaponAnimData(OrigenArma) = armaVacia
End If

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(WeaponAnimData)
    If WeaponAnimData(i).WeaponWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumArma Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveBota(ByVal NumBota As Integer, ByVal OrigenBota As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim botaVacia   As HeadData

If NumBota > UBound(BotasAnimData) Then ReDim Preserve BotasAnimData(0 To NumBota) As HeadData

If BotasAnimData(NumBota).Head(1).grhindex > 0 Then
    respuesta = MsgBox("La bota " & NumBota & " ya existe, ¿deseas sobreescribirla?", 4, "Aviso")
    If respuesta = vbYes Then
        BotasAnimData(NumBota) = BotasAnimData(OrigenBota)
        If BorrarOriginal Then BotasAnimData(OrigenBota) = botaVacia
    End If
Else
    BotasAnimData(NumBota) = BotasAnimData(OrigenBota)
    If BorrarOriginal Then BotasAnimData(OrigenBota) = botaVacia
End If


CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(BotasAnimData)
    If BotasAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumBota Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaBota(ByVal NumBota As Integer, Optional ByVal RefreshList As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumBota > UBound(BotasAnimData) Then ReDim Preserve BotasAnimData(0 To NumBota) As HeadData

BotasAnimData(NumBota).Head(1).grhindex = 1

If Not RefreshList Then Exit Sub
CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(BotasAnimData)
    If BotasAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumBota Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub AgregaCapa(ByVal NumCapa As Integer, Optional ByVal RefreshList As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If NumCapa > UBound(EspaldaAnimData) Then ReDim Preserve EspaldaAnimData(0 To NumCapa) As HeadData

EspaldaAnimData(NumCapa).Head(1).grhindex = 1

If Not RefreshList Then Exit Sub
CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(EspaldaAnimData)
    If EspaldaAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumCapa Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub AgregaFx(ByVal FxCapa As Integer)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long

If FxCapa > UBound(FxData) Then ReDim Preserve FxData(0 To FxCapa) As FxData

FxData(FxCapa).Fx.grhindex = 1

CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(FxData)
    If FxData(i).Fx.grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = FxCapa Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveCapa(ByVal NumCapa As Integer, ByVal origenCapa As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim botaVacia   As HeadData


If NumCapa > UBound(EspaldaAnimData) Then ReDim Preserve EspaldaAnimData(0 To NumCapa) As HeadData

If EspaldaAnimData(NumCapa).Head(1).grhindex > 0 Then
    respuesta = MsgBox("La capa " & NumCapa & " ya existe, ¿deseas sobreescribirla?", 4, "Aviso")
    If respuesta = vbYes Then
        EspaldaAnimData(NumCapa) = EspaldaAnimData(origenCapa)
        If BorrarOriginal Then EspaldaAnimData(origenCapa) = botaVacia
    End If
Else
    EspaldaAnimData(NumCapa) = EspaldaAnimData(origenCapa)
    If BorrarOriginal Then EspaldaAnimData(origenCapa) = botaVacia
End If


CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(EspaldaAnimData)
    If EspaldaAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumCapa Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub
Public Sub MueveFX(ByVal NumFx As Integer, ByVal origenFx As Integer, Optional ByVal BorrarOriginal As Boolean = True)
Dim i As Long
Dim EsteIndex As Long
Dim CuentaIndex As Long
Dim respuesta As Byte
Dim fxVacio   As FxData


If NumFx > UBound(FxData) Then ReDim Preserve FxData(0 To NumFx) As FxData

If FxData(NumFx).Fx.grhindex > 0 Then
    respuesta = MsgBox("El fx " & NumFx & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")
    If respuesta = vbYes Then
        FxData(NumFx) = FxData(origenFx)
        If BorrarOriginal Then FxData(origenFx) = fxVacio
    End If
Else
    FxData(NumFx) = FxData(origenFx)
    If BorrarOriginal Then FxData(origenFx) = fxVacio
End If


CuentaIndex = -1
frmMain.Lista.Clear
For i = 1 To UBound(FxData)
    If FxData(i).Fx.grhindex > 0 Then
        frmMain.Lista.AddItem i
        CuentaIndex = CuentaIndex + 1
    End If
    If i = NumFx Then
        EsteIndex = CuentaIndex
    End If
Next i

frmMain.Lista.listIndex = EsteIndex

End Sub

Public Sub RenuevaListaGrH()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To MAXGrH
    If GrhData(i).NumFrames = 1 Then
        frmMain.Lista.AddItem i
    ElseIf GrhData(i).NumFrames > 1 Then
        frmMain.Lista.AddItem i & " (animacion)"
    End If
Next i

End Sub
Public Sub RenuevaListaBodys()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(BodyData)
    If BodyData(i).Walk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaCabezas()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(HeadData)
    If HeadData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaCascos()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(CascoAnimData)
    If CascoAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaEscudos()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(ShieldAnimData)
    If ShieldAnimData(i).ShieldWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaArmas()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(WeaponAnimData)
    If WeaponAnimData(i).WeaponWalk(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaBotas()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(BotasAnimData)
    If BotasAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaCapas()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(EspaldaAnimData)
    If EspaldaAnimData(i).Head(1).grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub
Public Sub RenuevaListaFX()
Dim i As Long

frmMain.Lista.Clear

For i = 1 To UBound(FxData)
    If FxData(i).Fx.grhindex > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub

Public Sub RenuevaListaResource()
Dim i As Long


frmMain.Lista.Clear

For i = 1 To 32000
    If ExisteBMP(i) > 0 Then
        frmMain.Lista.AddItem i
    End If
Next i
End Sub

Public Function GrhCorrecto(ByRef GrhT As GrhData, ByRef ErrorMSG As String, ByRef ErroresGrh As ErroresGrh) As Long
' Comprueba que un grafico es correcto
Dim Alto As Long
Dim Ancho As Long
Dim i As Long
Dim DumyString As String
Dim PrimerAlto As Long
Dim PrimerAncho As Long
Dim dumyErroresGrh As ErroresGrh

ErroresGrh.ErrorCritico = False


If GrhT.NumFrames <= 0 Then
    ErrorMSG = "Nº de frames incorrecto"
    GrhCorrecto = 0
    ErroresGrh.ErrorCritico = True
    ErroresGrh.colores(2) = vbRed
    Exit Function
End If


If GrhT.NumFrames = 1 Then
    'si es solo un frame lo comprobamos
    GrhCorrecto = GrhCorrectoNormal(GrhT, ErrorMSG, ErroresGrh)
    ErroresGrh.EsAnimacion = False
Else
    ErroresGrh.EsAnimacion = True
    ' si es una animacion, comprobamos frame a frame
    For i = 1 To GrhT.NumFrames
        If GrhT.Frames(i) > 0 Then
            If GrhData(GrhT.Frames(i)).NumFrames <> 1 Or (GrhCorrectoNormal(GrhData(GrhT.Frames(i)), DumyString, dumyErroresGrh) < 2) Then
                ErrorMSG = ErrorMSG & "El frame nº " & i & " es incorrecto. "
                ErroresGrh.ErrorCritico = True
                GrhCorrecto = 1
                ErroresGrh.colores(1) = vbRed
            Else
                If i = 1 Then
                    PrimerAlto = GrhData(GrhT.Frames(i)).pixelHeight
                    PrimerAncho = GrhData(GrhT.Frames(i)).pixelWidth
                Else
                    Alto = GrhData(GrhT.Frames(i)).pixelHeight
                    Ancho = GrhData(GrhT.Frames(i)).pixelWidth
                    If Alto <> PrimerAlto Then
                        ErrorMSG = ErrorMSG & "El frame nº " & i & " distintas dimensiones. "
                        ErroresGrh.colores(1) = vbYellow
                    ElseIf Ancho <> PrimerAncho Then
                        ErrorMSG = ErrorMSG & "El frame nº " & i & " distintas dimensiones. "
                        ErroresGrh.colores(1) = vbYellow
                    End If
                End If
            End If
        Else
            ErrorMSG = ErrorMSG & "Falta frame nº " & i & ". "
            ErroresGrh.ErrorCritico = True
            ErroresGrh.colores(1) = vbRed
        End If
    Next i
End If


End Function

Public Function GrhCorrectoNormal(ByRef GrhT As GrhData, ByRef ErrorMSG As String, ByRef ErroresGrh As ErroresGrh) As Long
Dim Alto As Long
Dim Ancho As Long
Dim dumYin As Integer

'Comprueba que el grh es correcto. Ademas pone en rojo los texboxes con datos incorrectos.

    If GrhT.NumFrames <= 0 Then
        ErrorMSG = "Nº de frames incorrecto"
        GrhCorrectoNormal = 0
        ErroresGrh.colores(2) = vbRed
        ErroresGrh.ErrorCritico = True
        Exit Function
    End If
    
    If ExisteBMP(GrhT.FileNum) = ResourceFile Or (ResourceFile = 3 And ExisteBMP(GrhT.FileNum) > 0) Then
        Call GetTamañoBMP(GrhT.FileNum, Alto, Ancho, dumYin)
    Else
        ErrorMSG = "El archivo " & GrhT.FileNum & ".bmp no existe"
        GrhCorrectoNormal = 1
        ErroresGrh.colores(0) = vbRed
        ErroresGrh.ErrorCritico = True
        Exit Function
    End If
    
    GrhCorrectoNormal = 2 'mascara d bits, bit de grafico existente
    
    If GrhT.sX > Ancho Or GrhT.sY > Alto Then
        If GrhT.sX > Ancho Then
            ErrorMSG = ErrorMSG & "Posicion X fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 8 'mascara d bits , bit de error 2
            ErroresGrh.colores(6) = vbRed
        End If
        If GrhT.sY > Alto Then
            ErrorMSG = ErrorMSG & "Posicion Y fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 4 'mascara d bits , bit de error 1
            ErroresGrh.colores(7) = vbRed
        End If
    Else
        If GrhT.sY + GrhT.pixelHeight > Alto Then
            ErrorMSG = ErrorMSG & "Alto fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 16 'mascara d bits , bit de error 3
            ErroresGrh.colores(3) = vbYellow
        End If
        If GrhT.sX + GrhT.pixelWidth > Ancho Then
            ErrorMSG = ErrorMSG & "Ancho fuera del BMP. "
            GrhCorrectoNormal = GrhCorrectoNormal + 32 'mascara d bits , bit de error 4
            ErroresGrh.colores(4) = vbYellow
        End If
    End If
End Function
