Attribute VB_Name = "Module1"
Option Explicit

Private Type Obj
    Index As Integer
    Amount As Long
End Type

'Map
    Public Type MapHeader
        Name As String * 48
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
        dy As Long 'Dimensiones del Mapa Y (En Tiles)
        
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
        Color As Long
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
        NPCs() As tInfo
        particulas() As tInfo
        triggers() As tInfo
        
        exits() As tExits
        objs() As tObjs
        blocks() As tBlocks
        luces() As tLuces
    End Type
'Map

Public Sub Map_Load(ByVal map As Integer)
On Error Resume Next
    Dim MH As MapHeader
    Dim MAH As MapAfHeader
    Dim F As Integer
    Dim x As Long
    Dim y As Long
    Dim c1() As Integer
    Dim i As Long
    
    F = FreeFile
    Open App.path & "\Maps\mapa" & map & ".abr" For Binary Access Read As #F
        Get #F, , MH
        
        With MAH
            ReDim .blocks(MH.numBlocks)
            ReDim .capa2(MH.numCapa2)
            ReDim .capa3(MH.numCapa3)
            ReDim .capa4(MH.numCapa4)
            ReDim .objs(MH.numObjs)
            ReDim .NPCs(MH.numNpcs)
            ReDim .luces(MH.numLuces)
            ReDim .particulas(MH.numParticulas)
            ReDim .exits(MH.numExits)
            ReDim .triggers(MH.numTriggers)
        End With
        
        Get #F, , MAH
        
        ReDim c1(1 To MH.dX, 1 To MH.dy)
        Get #F, , c1
    Close #F

    If MH.dX = 0 Then MH.dX = 100
    If MH.dy = 0 Then MH.dy = 100

    ReDim MapData(1 To MH.dX, 1 To MH.dy)

    x = 1
    For i = 1 To MH.dX * MH.dy
        y = y + 1
        If y = MH.dy + 1 Then
            y = 1
            x = x + 1
        End If
        
        'InitGrh MapData(x, y).Graphic(1), c1(x, y)
   
        If MH.numCapa2 >= i Then
            InitGrh MapData(MAH.capa2(i).x, MAH.capa2(i).y).Graphic(2), MAH.capa2(i).info
        End If
        
        If MH.numCapa3 >= i Then
            InitGrh MapData(MAH.capa3(i).x, MAH.capa3(i).y).Graphic(3), MAH.capa3(i).info
        End If
        
        If MH.numCapa4 >= i Then
            InitGrh MapData(MAH.capa4(i).x, MAH.capa4(i).y).Graphic(4), MAH.capa4(i).info
        End If
        
        If MH.numBlocks >= i Then
            MapData(MAH.blocks(i).x, MAH.blocks(i).y).Blocked = True
        End If
        
        If MH.numTriggers >= i Then
            MapData(MAH.triggers(i).x, MAH.triggers(i).y).Trigger = MAH.triggers(i).info
        End If
    Next i

End Sub


