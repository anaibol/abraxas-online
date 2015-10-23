Attribute VB_Name = "Module1"
Option Explicit

'Map
    Public Type MapHeader
        Name As String * 48
        Music As Integer
        
        PK As Byte
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
        X As Integer
        Y As Integer
        info As Integer
    End Type
    
    Public Type tBlocks
        X As Integer
        Y As Integer
    End Type

    Public Type tExits
        X As Integer
        Y As Integer
        aex As WorldPos
    End Type
    
    Public Type tLuces
        X As Integer
        Y As Integer
        range As Byte
        color As Long
    End Type
    
    Public Type tObjs
        X As Integer
        Y As Integer
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

Public Sub Map_Load(ByVal Map As Integer)
On Error Resume Next
    Dim MH As MapHeader
    Dim MAH As MapAfHeader
    Dim f As Integer
    Dim X As Long
    Dim Y As Long
    Dim c1() As Integer
    Dim i As Long
    
    f = FreeFile
    Open App.Path & "\Maps\mapa" & Map & ".abr" For Binary Access Read As #f
        Get #f, , MH
        
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
        
        Get #f, , MAH
        
        ReDim c1(1 To MH.dX, 1 To MH.dY)
        Get #f, , c1
    Close #f

    If MH.dX = 0 Then MH.dX = 100
    If MH.dY = 0 Then MH.dY = 100

    ReDim MapData(1 To MH.dX, 1 To MH.dY)

    X = 1
    For i = 1 To MH.dX * MH.dY
        Y = Y + 1
        If Y = MH.dY + 1 Then
            Y = 1
            X = X + 1
        End If
        
        MapData(X, Y).Graphic(1) = c1(X, Y)
   
        If MH.numCapa2 >= i Then
            MapData(MAH.capa2(i).X, MAH.capa2(i).Y).Graphic(2) = MAH.capa2(i).info
        End If
        
        If MH.numCapa3 >= i Then
            MapData(MAH.capa3(i).X, MAH.capa3(i).Y).Graphic(3) = MAH.capa3(i).info
        End If
        
        If MH.numCapa4 >= i Then
            MapData(MAH.capa4(i).X, MAH.capa4(i).Y).Graphic(4) = MAH.capa4(i).info
        End If
        
        If MH.numBlocks >= i Then
            MapData(MAH.blocks(i).X, MAH.blocks(i).Y).Blocked = True
        End If
        
        If MH.numTriggers >= i Then
            MapData(MAH.triggers(i).X, MAH.triggers(i).Y).Trigger = MAH.triggers(i).info
        End If
        
        If MH.numObjs >= i Then
            If MAH.objs(i).info.index > 0 And MAH.objs(i).info.index < NumObjDatas Then
                MapData(MAH.objs(i).X, MAH.objs(i).Y).ObjInfo.index = MAH.objs(i).info.index
                MapData(MAH.objs(i).X, MAH.objs(i).Y).ObjInfo.Amount = MAH.objs(i).info.Amount
            End If
        End If
        
        If MH.numExits >= i Then
            MapData(MAH.exits(i).X, MAH.exits(i).Y).TileExit = MAH.exits(i).aex
        End If
        
        If MH.numNpcs >= i Then
        
        End If
    Next i
    
End Sub

