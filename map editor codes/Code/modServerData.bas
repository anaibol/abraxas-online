Attribute VB_Name = "modServerData"
Option Explicit

'Map
    Public Type MapHeader
        name As String * 48
        Music As Integer
        
        PK As Byte
        Poblacion As Integer
        notMagia As Byte
        notInvi As Byte
        notResu As Byte
        notRoboPermitido As Byte
        notInvocar As Byte
        notOcultar As Byte
        
        dX As Long 'Dimensiones del Mapa X (En Tiles)
        dY As Long 'Dimensiones del Mapa Y (En Tiles)
        
        terreno As Byte
        restringir As Byte
        
        numCapa2 As Integer
        numCapa3 As Integer
        numCapa4 As Integer
        numLuces As Integer
        numParticulas As Integer
        numBlocks As Integer
        numTriggers As Integer
        numNpcs As Integer
        numObjs As Integer
        numExits As Integer
    End Type
    
    Public Type tInfo
        x As Integer
        y As Integer
        info As Integer
    End Type
    
    Public Type tBlocks
        x As Integer
        y As Integer
    End Type

    Public Type tExits
        x As Integer
        y As Integer
        aex As WorldPos
    End Type
    
    Public Type tLuces
        x As Integer
        y As Integer
        range As Byte
        color As Long
    End Type
    
    Public Type tObjs
        x As Integer
        y As Integer
        info As Obj
    End Type
    
    Public Type MapAfHeader
        capa2() As tInfo
        capa3() As tInfo
        capa4() As tInfo
        npcs() As tInfo
        particulas() As tInfo
        triggers() As tInfo
        
        exits() As tExits
        objs() As tObjs
        blocks() As tBlocks
        luces() As tLuces
    End Type
'Map

Public Type NpcData
    name As String
    Body As Integer
    Head As Integer
    Heading As Byte
    Desc As String
    info As String
End Type
Public NumObjDatas As Integer

Public numNpcs As Integer
Public NpcData() As NpcData

Public Type ObjData
    name As String 'Nombre del obj
    
    OBJType As eObjType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
'    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    
    Mujer As Byte
    Hombre As Byte
    
    Envenena As Byte
    Paraliza As Byte
    
    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    Texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To 16) As eClass
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    
    NoSeCae As Integer
    
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    
    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
End Type

Public ObjData() As ObjData

Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo ErrHandler

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DirDat & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))

ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = Val(Leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = Val(Leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
        Case eObjType.otArmadura
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eObjType.otESCUDO
            ObjData(Object).ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eObjType.otCASCO
            ObjData(Object).CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eObjType.otWeapon
            ObjData(Object).WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = Val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = Val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = Val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eObjType.otInstrumentos
            ObjData(Object).Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
            'Pablo (ToxicWaste)
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eObjType.otMinerales
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eObjType.otPuertas, eObjType.otBotellaVacia, eObjType.otBotellaLlena
            ObjData(Object).IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eObjType.otBarcos
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eObjType.otFlechas
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
        Case eObjType.otAnillo 'Pablo (ToxicWaste)
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            
    End Select
    
    ObjData(Object).Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
    ObjData(Object).RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).RazaDrow = Val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
    ObjData(Object).RazaElfa = Val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
    ObjData(Object).RazaGnoma = Val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    ObjData(Object).RazaHumana = Val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
    ObjData(Object).Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).Texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")

    ObjData(Object).DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    frmMode.ObjList.AddItem ObjData(Object).name & " - #" & Object
Next Object

Set Leer = Nothing

Exit Sub

ErrHandler:
    MsgBox "error cargando objetos " & err.Number & ": " & err.Description


End Sub

Public Sub Map_Load_In2(ByVal map As String)
    Dim MH As MapHeader
    Dim MAH As MapAfHeader
    Dim F As Integer
    Dim x As Long
    Dim y As Long
    Dim c1(1 To 100, 1 To 100) As Integer
    Dim i As Long
    
    F = FreeFile
    Open map For Binary Access Read As #F
        Get #F, , MH
        
        With MAH
            ReDim .blocks(MH.numBlocks)
            ReDim .capa2(MH.numCapa2)
            ReDim .capa3(MH.numCapa3)
            ReDim .capa4(MH.numCapa4)
            ReDim .objs(MH.numObjs)
            ReDim .npcs(MH.numNpcs)
            ReDim .luces(MH.numLuces)
            ReDim .particulas(MH.numParticulas)
            ReDim .exits(MH.numExits)
            ReDim .triggers(MH.numTriggers)
        End With
        
        Get #F, , MAH
        Get #F, , c1
    Close #F

    ReDim MapData2(1 To 100, 1 To 100)
    
    x = 1
    For i = 1 To 10000
        y = y + 1
        If y = 101 Then
            y = 1
            x = x + 1
        End If
        
        Grh_Init MapData2(x, y).Graphic(1), c1(x, y)
   
        If MH.numCapa2 >= i Then
            Grh_Init MapData2(MAH.capa2(i).x, MAH.capa2(i).y).Graphic(2), MAH.capa2(i).info
        End If
        
        If MH.numCapa3 >= i Then
            Grh_Init MapData2(MAH.capa3(i).x, MAH.capa3(i).y).Graphic(3), MAH.capa3(i).info
        End If
        
        If MH.numCapa4 >= i Then
            Grh_Init MapData2(MAH.capa4(i).x, MAH.capa4(i).y).Graphic(4), MAH.capa4(i).info
        End If
        
        If MH.numBlocks >= i Then
            MapData2(MAH.blocks(i).x, MAH.blocks(i).y).Blocked = 1
        End If
        
        If MH.numTriggers >= i Then
            MapData2(MAH.triggers(i).x, MAH.triggers(i).y).Trigger = MAH.triggers(i).info
        End If
        
        If MH.numNpcs >= i Then
            MapData2(MAH.npcs(i).x, MAH.npcs(i).y).NPCIndex = MAH.npcs(i).info
        End If
    Next i
End Sub

'Formato viejo de abraxas
Public Sub Map_Load_Old(ByVal map As String)
On Error Resume Next

    Dim mapString As String
    Dim hMap As Integer
    Dim hInf As Integer
    Dim tempInt As Integer
    Dim x As Long, y As Long
    Dim ByFlags As Byte
    Dim Body As Integer, Head As Integer, Heading As Integer
    
    mapString = map
    
    'Mannakia : Deduzco que la carga es asi.
    hMap = FreeFile()
    Open mapString & ".map" For Binary As #hMap
        Seek hMap, 1

    hInf = FreeFile()
    Open mapString & ".inf" For Binary As #hInf
        Seek hMap, 268 ' 1 + (255 + 4 + 4 = 263 -> tCabecera) + (4 = getDouble)
        Seek hInf, 7 ' 1 + (6 = getDouble + getInteger)

        For y = 1 To 100
            For x = 1 To 100
                With MapData(x, y)
                    '.map file
                    Get hMap, , ByFlags
                
                    If ByFlags And 1 Then .Blocked = 1
                    
                    Get hMap, , tempInt
                    Grh_Init MapData(x, y).Graphic(1), tempInt
                    
                    'Layer 2 used?
                    If ByFlags And 2 Then
                        Get hMap, , tempInt
                        Grh_Init MapData(x, y).Graphic(2), tempInt
                    End If
    
                    'Layer 3 used?
                    If ByFlags And 4 Then
                        Get hMap, , tempInt
                        Grh_Init MapData(x, y).Graphic(3), tempInt
                    End If
    
                    'Layer 4 used?
                    If ByFlags And 8 Then
                        Get hMap, , tempInt
                        Grh_Init MapData(x, y).Graphic(4), tempInt
                    End If
                    
                    'Trigger used?
                    If ByFlags And 16 Then
                        Get hMap, , tempInt
                        .Trigger = tempInt
                    End If
                    
                    '.inf file
                    Get hInf, , ByFlags
                    
                    'Mannakia : ¿?
                    'If ByFlags And 1 Then
                    '    .TileExit.Map = InfReader.getInteger
                    '    .TileExit.X = InfReader.getInteger
                    '    .TileExit.Y = InfReader.getInteger
                    'End If
                    
                    If ByFlags And 1 Then
                        Get hInf, , tempInt
                        Get hInf, , tempInt
                        Get hInf, , tempInt
                    End If
                    
                    If ByFlags And 2 Then
                        Get hInf, , tempInt
                        .NPCIndex = tempInt
                    
                        Body = NpcData(tempInt).Body
                        Head = NpcData(tempInt).Head
                        Heading = NpcData(tempInt).Heading
                        
                        Engine.Char_Make NextOpenChar, Body, Head, Heading, x, y, 2, 2, 2
                    End If
    
                    If ByFlags And 4 Then
                        'Get and make Object
                        Get hInf, , .OBJInfo.ObjIndex
                        Get hInf, , .OBJInfo.amount
                        
                        Grh_Init MapData(x, y).ObjGrh, ObjData(.OBJInfo.ObjIndex).GrhIndex
                    End If
                End With
            Next x
        Next y
    Close hMap
    Close hInf
End Sub

Public Sub Map_Load(ByVal map As String)
On Error Resume Next
    Dim MH As MapHeader
    Dim MAH As MapAfHeader
    Dim F As Integer
    Dim x As Long
    Dim y As Long
    Dim c1() As Integer
    Dim i As Long
    
    F = FreeFile
    Open map For Binary Access Read As #F
        Get #F, , MH
        
        With MAH
            ReDim .blocks(MH.numBlocks)
            ReDim .capa2(MH.numCapa2)
            ReDim .capa3(MH.numCapa3)
            ReDim .capa4(MH.numCapa4)
            ReDim .objs(MH.numObjs)
            ReDim .npcs(MH.numNpcs)
            ReDim .luces(MH.numLuces)
            ReDim .particulas(MH.numParticulas)
            ReDim .exits(MH.numExits)
            ReDim .triggers(MH.numTriggers)
        End With
        
        Get #F, , MAH
        
        ReDim c1(1 To MH.dX, 1 To MH.dY)
        Get #F, , c1
    Close #F

    If MH.dX = 0 Then MH.dX = 100
    If MH.dY = 0 Then MH.dY = 100

    ReDim MapData(1 To MH.dX, 1 To MH.dY)
    
    frmMinimap.MiniMap.Cls
    frmMinimap.MiniMap.BackColor = 0
    
    x = 1
    For i = 1 To MH.dX * MH.dY
        y = y + 1
        If y = MH.dY + 1 Then
            y = 1
            x = x + 1
        End If
        
        Grh_Init MapData(x, y).Graphic(1), c1(x, y)
   
        If MH.numCapa2 >= i Then
            Grh_Init MapData(MAH.capa2(i).x, MAH.capa2(i).y).Graphic(2), MAH.capa2(i).info
        End If
        
        If MH.numCapa3 >= i Then
            Grh_Init MapData(MAH.capa3(i).x, MAH.capa3(i).y).Graphic(3), MAH.capa3(i).info
        End If
        
        If MH.numCapa4 >= i Then
            Grh_Init MapData(MAH.capa4(i).x, MAH.capa4(i).y).Graphic(4), MAH.capa4(i).info
        End If
        
        If MH.numBlocks >= i Then
            MapData(MAH.blocks(i).x, MAH.blocks(i).y).Blocked = 1
        End If
        
        If MH.numTriggers >= i Then
            MapData(MAH.triggers(i).x, MAH.triggers(i).y).Trigger = MAH.triggers(i).info
        End If
        
        If MH.numParticulas >= i Then
            MapData(MAH.particulas(i).x, MAH.particulas(i).y).Particle_index = MAH.particulas(i).info
            Engine.Particle_Save_Create MAH.particulas(i).info, MapData(MAH.particulas(i).x, MAH.particulas(i).y).Particle_Group
        End If
        
        If MH.numLuces >= i Then
            Engine.Light_Create MAH.luces(i).x, MAH.luces(i).y, MAH.luces(i).range, MAH.luces(i).color
        End If
        
        If MH.numObjs >= i Then
            If MAH.objs(i).info.ObjIndex > 0 And MAH.objs(i).info.ObjIndex < NumObjDatas Then
                MapData(MAH.objs(i).x, MAH.objs(i).y).OBJInfo.ObjIndex = MAH.objs(i).info.ObjIndex
                MapData(MAH.objs(i).x, MAH.objs(i).y).OBJInfo.amount = MAH.objs(i).info.amount
            
                Grh_Init MapData(MAH.objs(i).x, MAH.objs(i).y).ObjGrh, ObjData(MAH.objs(i).info.ObjIndex).GrhIndex
            End If
        End If
        
        If MH.numExits >= i Then
            MapData(MAH.exits(i).x, MAH.exits(i).y).TileExit = MAH.exits(i).aex
        End If
        
        If MH.numNpcs >= i Then
            MapData(MAH.npcs(i).x, MAH.npcs(i).y).NPCIndex = MAH.npcs(i).info
            Dim Body As Integer, Head As Integer, Heading As Integer
            Body = NpcData(MAH.npcs(i).info).Body
            Head = NpcData(MAH.npcs(i).info).Head
            Heading = NpcData(MAH.npcs(i).info).Heading
            
            Engine.Char_Make NextOpenChar, Body, Head, Heading, MAH.npcs(i).x, MAH.npcs(i).y, 2, 2, 2
        End If
    Next i
    
    With MapInfo
        .PK = MH.PK
        .InviSinEfecto = MH.notInvi
        .MagiaSinEfecto = MH.notMagia
        .restringir = MH.notResu
        .Music = MH.Music
        .name = RTrim$(MH.name)
        .restringir = MH.restringir
        .terreno = MH.terreno
        .dX = MH.dX
        .dY = MH.dY
    End With
    
    Cargar_Mode_Map
    
    MinXBorder = (1 + (frmMain.ScaleWidth / 64)) - 1
    MaxXBorder = (MapInfo.dX - (frmMain.ScaleWidth / 64)) + 1
    MinYBorder = (1 + (frmMain.ScaleHeight / 64)) + 1
    MaxYBorder = (MapInfo.dY - (frmMain.ScaleHeight / 64)) - 1
    
    frmMain.FileMap = map
    DameArchivo map, map, frmMain.FileMapDir
End Sub

Public Sub Map_Save(ByVal route As String)
    Dim MH As MapHeader
    Dim MAH As MapAfHeader
    
    With MAH
        ReDim .blocks(10000)
        ReDim .capa2(10000)
        ReDim .capa3(10000)
        ReDim .capa4(10000)
        ReDim .objs(10000)
        ReDim .npcs(10000)
        ReDim .luces(10000)
        ReDim .particulas(10000)
        ReDim .exits(10000)
        ReDim .triggers(10000)
    End With
    
    Dim x As Long
    Dim y As Long
    Dim c1() As Integer
    
    MH.dX = MapInfo.dX
    MH.dY = MapInfo.dY
    
    ReDim c1(1 To MH.dX, 1 To MH.dY) As Integer
    
    For x = 1 To MH.dX
        For y = 1 To MH.dY
            c1(x, y) = MapData(x, y).Graphic(1).GrhIndex
            
            If MapData(x, y).Graphic(2).GrhIndex > 0 Then
                MH.numCapa2 = MH.numCapa2 + 1
                MAH.capa2(MH.numCapa2).x = x
                MAH.capa2(MH.numCapa2).y = y
                MAH.capa2(MH.numCapa2).info = MapData(x, y).Graphic(2).GrhIndex
            End If
            
            If MapData(x, y).Graphic(3).GrhIndex > 0 Then
                MH.numCapa3 = MH.numCapa3 + 1
                MAH.capa3(MH.numCapa3).x = x
                MAH.capa3(MH.numCapa3).y = y
                MAH.capa3(MH.numCapa3).info = MapData(x, y).Graphic(3).GrhIndex
            End If
            
            If MapData(x, y).Graphic(4).GrhIndex > 0 Then
                MH.numCapa4 = MH.numCapa4 + 1
                MAH.capa4(MH.numCapa4).x = x
                MAH.capa4(MH.numCapa4).y = y
                MAH.capa4(MH.numCapa4).info = MapData(x, y).Graphic(4).GrhIndex
            End If
            
            If MapData(x, y).Blocked > 0 Then
                MH.numBlocks = MH.numBlocks + 1
                MAH.blocks(MH.numBlocks).x = x
                MAH.blocks(MH.numBlocks).y = y
            End If
            
            If MapData(x, y).NPCIndex > 0 Then
                MH.numNpcs = MH.numNpcs + 1
                MAH.npcs(MH.numNpcs).x = x
                MAH.npcs(MH.numNpcs).y = y
                MAH.npcs(MH.numNpcs).info = MapData(x, y).NPCIndex
            End If
            
            If MapData(x, y).OBJInfo.ObjIndex > 0 Then
                MH.numObjs = MH.numObjs + 1
                MAH.objs(MH.numObjs).x = x
                MAH.objs(MH.numObjs).y = y
                MAH.objs(MH.numObjs).info.ObjIndex = MapData(x, y).OBJInfo.ObjIndex
                MAH.objs(MH.numObjs).info.amount = MapData(x, y).OBJInfo.amount
            End If
            
            If MapData(x, y).TileExit.map > 0 Then
                MH.numExits = MH.numExits + 1
                MAH.exits(MH.numExits).x = x
                MAH.exits(MH.numExits).y = y
                MAH.exits(MH.numExits).aex.map = MapData(x, y).TileExit.map
                MAH.exits(MH.numExits).aex.x = MapData(x, y).TileExit.x
                MAH.exits(MH.numExits).aex.y = MapData(x, y).TileExit.y
            End If
            
            If MapData(x, y).Particle_index > 0 Then
                MH.numParticulas = MH.numParticulas + 1
                MAH.particulas(MH.numParticulas).x = x
                MAH.particulas(MH.numParticulas).y = y
                MAH.particulas(MH.numParticulas).info = MapData(x, y).Particle_index
            End If
            
            If MapData(x, y).Trigger > 0 Then
                MH.numTriggers = MH.numTriggers + 1
                MAH.triggers(MH.numTriggers).x = x
                MAH.triggers(MH.numTriggers).y = y
                MAH.triggers(MH.numTriggers).info = MapData(x, y).Trigger
            End If
            
            If MapData(x, y).Luz <> 0 Then
                MH.numLuces = MH.numLuces + 1
                MAH.luces(MH.numLuces).x = x
                MAH.luces(MH.numLuces).y = y
                MAH.luces(MH.numLuces).color = light_list(MapData(x, y).Luz).color
                MAH.luces(MH.numLuces).range = light_list(MapData(x, y).Luz).range
            End If
        Next y
    Next x
        
    With MAH
        ReDim Preserve .blocks(MH.numBlocks)
        ReDim Preserve .capa2(MH.numCapa2)
        ReDim Preserve .capa3(MH.numCapa3)
        ReDim Preserve .capa4(MH.numCapa4)
        ReDim Preserve .objs(MH.numObjs)
        ReDim Preserve .npcs(MH.numNpcs)
        ReDim Preserve .luces(MH.numLuces)
        ReDim Preserve .particulas(MH.numParticulas)
        ReDim Preserve .exits(MH.numExits)
        ReDim Preserve .triggers(MH.numTriggers)
    End With
    
    With MH
        .name = MapInfo.name
        .Music = Val(MapInfo.Music)
        .notInvi = MapInfo.InviSinEfecto
        .notMagia = MapInfo.MagiaSinEfecto
        .notResu = MapInfo.ResuSinEfecto
        .PK = IIf(MapInfo.PK, 1, 0)
        .restringir = MapInfo.restringir
        .terreno = MapInfo.terreno
        .dX = MapInfo.dX
        .dY = MapInfo.dY
    End With
    
    Do While LCase$(Right$(route, 4)) = ".abr"
        route = Left$(route, Len(route) - 4)
    Loop
    
    route = LCase$(route) & ".abr"
    If FileExist(route, vbNormal) Then Kill route

    Open route For Binary Access Write As #1
        Put #1, , MH
        Put #1, , MAH
        Put #1, , c1
    Close #1
End Sub


Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

On Error GoTo Fallo

    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DirDat & "NPCs.dat")
    
    numNpcs = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    ReDim NpcData(1 To numNpcs) As NpcData
    
    Trabajando = "Dat\NPCs.dat"
    
    Call Leer.Initialize(DirDat & "NPCs.dat")
    
    For NPC = 1 To numNpcs
        NpcData(NPC).name = Leer.GetValue("NPC" & NPC, "Name")
        NpcData(NPC).Desc = Leer.GetValue("NPC" & NPC, "Desc")
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        NpcData(NPC).info = Leer.GetValue("NPC" & NPC, "Info")
        frmMode.lstNPC.AddItem NPC & "^" & NpcData(NPC).name
    Next NPC

    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDat & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

Public Sub Map_New()
    Erase MapData
    
    Dim dX As Long, dY As Long
    dX = Val(InputBox("Ingrese las dimensiones en X del mapa (En Tiles)", , 100))
    dY = Val(InputBox("Ingrese las dimensiones en Y del mapa (En Tiles)", , 100))
    
    MapInfo.dX = dX
    MapInfo.dY = dY
    
    MinXBorder = (1 + (frmMain.ScaleWidth / 64)) - 1
    MaxXBorder = (MapInfo.dX - (frmMain.ScaleWidth / 64)) + 1
    MinYBorder = (1 + (frmMain.ScaleHeight / 64)) + 1
    MaxYBorder = (MapInfo.dY - (frmMain.ScaleHeight / 64)) - 1
    
    ReDim MapData(1 To MapInfo.dX, 1 To MapInfo.dY) As MapBlock
End Sub

Public Sub Cargar_Mode_Map()
    If MapInfo.PK = True Then
        frmMode.PK.value = vbChecked
    Else
        frmMode.PK.value = vbUnchecked
    End If
    
    If MapInfo.ResuSinEfecto = 1 Then
        frmMode.RESU.value = vbChecked
    Else
        frmMode.RESU.value = vbUnchecked
    End If
    
    If MapInfo.InviSinEfecto = 1 Then
        frmMode.INVI.value = vbChecked
    Else
        frmMode.INVI.value = vbUnchecked
    End If
    
    If MapInfo.MagiaSinEfecto = 1 Then
        frmMode.Magia.value = vbChecked
    Else
        frmMode.Magia.value = vbUnchecked
    End If

    frmMode.txtNameMap.Text = MapInfo.name
    frmMode.txtMusicNum.Text = MapInfo.Music
    
    frmMode.DimMap = "Dimensiones del Mapa : X = " & MapInfo.dX & " ; Y =" & MapInfo.dY
    
    frmMode.terreno.ListIndex = MapInfo.terreno
    frmMode.restringuir.ListIndex = MapInfo.restringir
End Sub
