Attribute VB_Name = "PathFinding"
Option Explicit

Private Const ROWS As Integer = 500
Private Const COLUMS As Integer = 300
Private Const MaxINT As Integer = 1000

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
    Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal Row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
    IsWalkable = Not MapData(Row, Col).Blocked And MapData(Row, Col).NpcIndex < 1
    
    If MapData(Row, Col).UserIndex > 0 Then
        If MapData(Row, Col).UserIndex <> NpcList(NpcIndex).TargetUser Then
            IsWalkable = False
        End If
    End If
End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
    
    Dim V As tVertice
    Dim j As Integer
    
    'Look to North
    j = vfila - 1
    
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            'Nos aseguramos que no hay un camino más corto
            If T(j, vcolu).DistV = MaxINT Then
                'Actualizamos la tabla de calculos intermedios
                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.X = vcolu
                T(j, vcolu).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.X = vcolu
                V.Y = j
                Call Push(V)
            End If
        End If
    End If
    
    j = vfila + 1
    
    'look to south
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            'Nos aseguramos que no hay un camino más corto
            If T(j, vcolu).DistV = MaxINT Then
                'Actualizamos la tabla de calculos intermedios
                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.X = vcolu
                T(j, vcolu).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.X = vcolu
                V.Y = j
                Call Push(V)
            End If
        End If
    End If
    
    'look to west
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(vfila, vcolu - 1).DistV = MaxINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
                    T(vfila, vcolu - 1).PrevV.X = vcolu
                    T(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu - 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
    
    'look to east
    If Limites(vfila, vcolu + 1) Then
        If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
            'Nos aseguramos que no hay un camino más corto
            If T(vfila, vcolu + 1).DistV = MaxINT Then
                'Actualizamos la tabla de calculos intermedios
                T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
                T(vfila, vcolu + 1).PrevV.X = vcolu
                T(vfila, vcolu + 1).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.X = vcolu + 1
                V.Y = vfila
                Call Push(V)
            End If
        End If
    End If
    
End Sub

Public Sub SeekPath(ByVal NpcIndex As Integer)
'This PUBLIC SUB seeks a path from the NpcList(NpcIndex).pos
'to the location NpcList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the Maximum of steps
'allowed for the path.

    Dim cur_npc_pos As tVertice
    Dim tar_npc_pos As tVertice
    Dim V As tVertice
    Dim NpcMap As Integer
    
    NpcMap = NpcList(NpcIndex).Pos.Map
        
    cur_npc_pos.X = NpcList(NpcIndex).Pos.X
    cur_npc_pos.Y = NpcList(NpcIndex).Pos.Y
    
    tar_npc_pos.X = NpcList(NpcIndex).PFINFO.Target.X
    tar_npc_pos.Y = NpcList(NpcIndex).PFINFO.Target.Y
    
    Call InitializeTable(TmpArray, cur_npc_pos)
    
    Call InitQueue
    
    'We add the first vertex to the Queue
    Call Push(cur_npc_pos)
    
    Do While (Not IsEmpty)
        V = Pop
                
        If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then
            Exit Do
        End If
        
        Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
    Loop
    
    Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'Builds the path previously calculated

    Dim Pasos As Integer
    Dim miV As tVertice
    Dim i As Integer
    
    With NpcList(NpcIndex)
        Pasos = TmpArray(.PFINFO.Target.Y, .PFINFO.Target.X).DistV

        If Pasos = MaxINT Or Pasos = 0 Then
            .PFINFO.PathLenght = 0
            Exit Sub
        Else
            .PFINFO.PathLenght = Pasos
        End If
        
        If .PFINFO.PathLenght > 7 Then
            .PFINFO.PathLenght = 7
        End If
        
        ReDim .PFINFO.Path(0 To Pasos) As tVertice
        
        miV.X = .PFINFO.Target.X
        miV.Y = .PFINFO.Target.Y
        
        For i = Pasos To 1 Step -1
            .PFINFO.Path(i) = miV
            miV = TmpArray(miV.Y, miV.X).PrevV
        Next i
        
        .PFINFO.CurPos = 1
    End With
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice)
'Initialize the array where we calculate the path

    Dim MaxSteps As Byte
    
    MaxSteps = 7

    Dim j As Integer, k As Integer
    
    For j = S.Y - MaxSteps To S.Y + MaxSteps
        For k = S.X - MaxSteps To S.X + MaxSteps
            If InMapBounds(1, j, k) Then
                T(j, k).Known = False
                T(j, k).DistV = MaxINT
                T(j, k).PrevV.X = 0
                T(j, k).PrevV.Y = 0
            End If
        Next
    Next
    
    T(S.Y, S.X).Known = False
    T(S.Y, S.X).DistV = 0
    
End Sub
