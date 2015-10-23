VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuadTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Quadtree structure - Independant QuadTree
' Each branch only descends to the required level
' Not all branches descend to the "lowest" level
' "Lowest" level is determined by "Number of Objects" in Cell (not by 'Cell Size')
' Each Cell (except the root cell) holds an array of all the objects beneath
'           This speeds up Viewport by A LOT (x22) But does require a "large" memory overhead.
'           Worst case (NObjects * Max Level) * 4 bytes (could be an Mb or 2)

' Quad Cell Locations
' /-----------\  0 - Top Left
' |  0  |  1  |  1 - Top Right
' |-----+-----|  2 - Bottom Left
' |  2  |  3  |  3 - Bottom Roight
' \-----------/

Private mTop As Boolean         ' Lets Instance know if it's the Root
Private mReferences() As Long   ' Private Array to Hold references for this cell
Private mnReferences As Long    ' Private Count of References
Private mSubdivided As Boolean  ' Indicates whether this cell contains sub-divisions
Private Subdivisions() As clsQuadTree   ' Array holds the sub-division members (if needed)

' "Real World" coordinates of the cell.
Private mbl As coord            ' Bottom Left of cell
Private mtr As coord            ' Top Right of Cell
Private mcentre As coord        ' Centre of Cell (Calc'd the Stored for Point Test)
Private mPolyArea As Double     ' the Area of the Root Cell

Public Property Let IsTop(ByVal b As Boolean)
    mTop = b
End Property

Public Sub CreateTree(ByRef References() As Long, ByVal nReferences As Long, ByVal blx As Double, ByVal bly As Double, ByVal trx As Double, ByVal try As Double)

' References() - Array of Object References
' nReferences  - Number of Elements Used in Array  (quicker to pass it than Redim Preserve, then Ubound)
' blx - Bottom Left "Real World" X Value of Cell
' bly - Bottom Left "Real World" Y Value of Cell
' trx - Top Right "Real World" X Value of Cell
' try - Top Right "Real World" Y Value of Cell
    
    Dim i As Long, j As Long
    Dim refs() As Long, nRefs As Long
    ' Store Cell coords
    mbl.X = blx
    mbl.Y = bly
    mtr.X = trx
    mtr.Y = try
    
    If mTop Then ' If root cell then
        maxReference = nReferences  ' Store Number of Objects
        ' Fixed Subdivide
        ' noSubdivide = 10
        
        ' Adaptive Subdivide, the more objects the larger nosubdivide is.
        noSubDivide = 10 * Int(Log(maxReference) / Log(10))
        
        noSubDivide = 10
        
        mPolyArea = (mtr.X - mbl.X) * (mtr.Y - mbl.Y)  ' Calculate the Area of the Root Cell
    Else  ' If not ROOT cell
        If nReferences = -1 Then
            mnReferences = 0
        Else
            ReDim mReferences(nReferences)   ' redim array
            mnReferences = nReferences + 1   ' and store
            For i = 0 To nReferences         ' the object references for this cell
                mReferences(i) = References(i)
            Next i
        End If
    End If
    
    If nReferences < noSubDivide Then
        mSubdivided = False              ' We don't need to sub-divide (we're done here)
    Else
        mcentre.X = (mbl.X + mtr.X) / 2  ' Calculate Centre.x
        mcentre.Y = (mbl.Y + mtr.Y) / 2  ' Calculate Centre.y
        mSubdivided = True               ' Mark this member as subdivided
        ReDim Subdivisions(3)            ' Redim Array top hold subdivisions
        ReDim refs(nReferences)          ' Redim Refs() to hold Object References
        For i = 0 To 3                   ' For Each Quadrant
            Set Subdivisions(i) = New clsQuadTree   ' Create new member
            nRefs = 0                    ' Set Object Counter to 0
            
            ' Ifs Faster than Select Case
            If i = 0 Then
                For j = 0 To nReferences ' Test Each Object
                   If PolyRects(References(j)).tr.X < mbl.X Then
                   ElseIf PolyRects(References(j)).bl.X >= mcentre.X Then
                   ElseIf PolyRects(References(j)).tr.Y < mcentre.Y Then
                   ElseIf PolyRects(References(j)).bl.Y >= mtr.Y Then
                   Else ' if it touches this quadrant (0) add it to the array
                      refs(nRefs) = References(j)
                      nRefs = nRefs + 1
                   End If
                Next j
                ' Call Create tree in quadrant 0, passing the array, nelements (nrefs-1) and "real world" coords
                Subdivisions(0).CreateTree refs(), nRefs - 1, mbl.X, mcentre.Y, mcentre.X, mtr.Y
            ElseIf i = 1 Then
                For j = 0 To nReferences
                   If PolyRects(References(j)).tr.X < mcentre.X Then
                   ElseIf PolyRects(References(j)).bl.X >= mtr.X Then
                   ElseIf PolyRects(References(j)).tr.Y < mcentre.Y Then
                   ElseIf PolyRects(References(j)).bl.Y >= mtr.Y Then
                   Else ' the object touches quadrant 1
                      refs(nRefs) = References(j) ' so add it to the array
                      nRefs = nRefs + 1
                   End If
                Next j
                Subdivisions(1).CreateTree refs(), nRefs - 1, mcentre.X, mcentre.Y, mtr.X, mtr.Y
            ElseIf i = 2 Then
                For j = 0 To nReferences
                   If PolyRects(References(j)).tr.X < mbl.X Then
                   ElseIf PolyRects(References(j)).bl.X >= mcentre.X Then
                   ElseIf PolyRects(References(j)).tr.Y < mbl.Y Then
                   ElseIf PolyRects(References(j)).bl.Y >= mcentre.Y Then
                   Else ' the object touches quadrant 2
                      refs(nRefs) = References(j) ' so add it to the array
                      nRefs = nRefs + 1
                   End If
                Next j
                Subdivisions(2).CreateTree refs(), nRefs - 1, mbl.X, mbl.Y, mcentre.X, mcentre.Y
            Else ' i was not 0,1 or 2 (MUST be 3)
                For j = 0 To nReferences
                   If PolyRects(References(j)).tr.X < mcentre.X Then
                   ElseIf PolyRects(References(j)).bl.X >= mtr.X Then
                   ElseIf PolyRects(References(j)).tr.Y < mbl.Y Then
                   ElseIf PolyRects(References(j)).bl.Y >= mcentre.Y Then
                   Else ' the object touches quadrant 3
                      refs(nRefs) = References(j) ' so add it to the array
                      nRefs = nRefs + 1
                   End If
                Next j
                Subdivisions(3).CreateTree refs(), nRefs - 1, mcentre.X, mbl.Y, mtr.X, mcentre.Y
            End If
        Next i
    End If
    
End Sub

Public Function PointTest() As Boolean

    ' Returns True if the cell hit contained object(s),
    ' False if the cell was empty
    If mTop Then
        ' initialise the results structures We know we'll only be getting 1 cell
        ' so init output for whatever the max objects in a cell is. (nosubdivide)
        ReDim QuadOutput(noSubDivide)
        nQuadOutput = 0
        ubQuadOutput = noSubDivide
    End If

    If mSubdivided Then ' If this cell is subdivded
        If pt.X < mcentre.X Then  ' are we in quadrants 0 or 2 - Left
            If pt.Y < mcentre.Y Then ' are in the bottom half
                PointTest = Subdivisions(2).PointTest()
            Else
                PointTest = Subdivisions(0).PointTest()
            End If
        Else ' not on left, must be on right (1 or 3)
            If pt.Y < mcentre.Y Then  ' are we in the bottom half
                PointTest = Subdivisions(3).PointTest()
            Else
                PointTest = Subdivisions(1).PointTest()
            End If
        End If
    Else ' If the cell is not subdivided
        If mnReferences > 0 Then ' if it contains objects
            PointTest = True     ' return true and
            Dim i As Long
            For i = 0 To mnReferences - 1  ' put the object references in the output array
                QuadOutput(i) = mReferences(i)
            Next i
            nQuadOutput = mnReferences
        End If
    End If
            
End Function

Public Sub ViewPort()
    ' Overview
    ' Returns in the QuadOutput() array references to all the objects that lie within
    ' a viewport, defined by the vp Rectangle structure.
    
    ' Details
    ' When called, recursively descends thru the quad tree.
    ' Locating Cells that lie partially or wholey within the viewport area (vp)
    ' If the current cell lies wholey within the vp then its entire contents are added to the output array
    ' If the cell lies partially within the vp the routine either
    '    a) moves "recursively" to testing the subdivisions (if the cell is subdivided)
    '    b) Returns all object references (if the cell is NOT subdivided)
    
    Dim i As Long
    ' call a function to do all the stuff we need
    ' when this is first called, setting up output arrays..
    ' and a few initial quick tests
    If mTop Then
        If Viewport_Top() Then Exit Sub
    End If
    
    ' does viewport touch cell ?
    If vp.tr.X < mbl.X Then   ' Right Edge of Viewport is Left of Left Edge of Cell
    ElseIf vp.bl.X > mtr.X Then ' Left Edge of Viewport is Right of Right Edge of Cell
    ElseIf vp.tr.Y < mbl.Y Then ' Top Edge of Viewport is Below of Bottom Edge of Cell
    ElseIf vp.bl.Y > mtr.Y Then ' Bottom Edge of Viewport is Above of Top Edge of Cell
    Else ' Viewport MUST (does) touch cell
        If mSubdivided Then
            ' does viewport completely enclose cell
            If vp.bl.X <= mbl.X Then  ' Left Edge of Viewport is Left of Left Edge of Cell
               If vp.tr.X >= mtr.X Then ' Right Edge of Viewport is Right of Right Edge of Cell
                   If vp.bl.Y <= mbl.Y Then ' Bottom Edge of Viewport is Below Bottom Edge of Cell
                      If vp.tr.Y >= mtr.Y Then ' Top Edge of Viewport is Above Top Edge of Cell
                          ' yes viewport encloses whole cell.... so we don't need to test anymore things,
                          ' just go and get all the references below here.
                          RetrieveData
                          Exit Sub
                        End If
                    End If
                End If
            End If
            ' we can't take the whole of this cell.. so descend a level
            ' Invoke Viewport Method of the sub-Cells
            For i = 0 To 3
                 Subdivisions(i).ViewPort
            Next i
        Else
            ' If this is the lowest cell level then return the data
            RetrieveData
        End If
    End If

End Sub

Private Function RetrieveData()

        Dim i As Long
        
        ' check if the output array is large enough to hold the new data
        If nQuadOutput + mnReferences > ubQuadOutput Then
            ' if not - increase the size of the output array
            ubQuadOutput = ubQuadOutput + mnReferences + AddRedim
            ReDim Preserve QuadOutput(ubQuadOutput)
        End If
        
        ' put all our references in QuadOutput
        For i = 0 To mnReferences - 1
            ' check if this reference has been returned before (during this viewport call)
            If blnRef(mReferences(i)) = 0 Then
                ' if not add the object reference to the output array
                QuadOutput(nQuadOutput) = mReferences(i)
                nQuadOutput = nQuadOutput + 1
                ' add store the fact that this refernce has been returned
                blnRef(mReferences(i)) = blnRef(mReferences(i)) + 1
             End If
        Next i
        
End Function
Private Sub Class_Terminate()
    Dim i As Long
    If mSubdivided Then ' if this member have been subdivided
        For i = 0 To 3  ' terminate its sub-members
            Set Subdivisions(i) = Nothing
        Next i
    End If
End Sub

Private Function Viewport_Top() As Boolean

        ' initialise the results structures
        Dim i As Long, areafrac As Double
        Dim X(1) As Double, Y(1) As Double, intarea As Double
    
        ' calculate the fraction of the Root Cell that COULD BE (not IS) covered
        areafrac = ((vp.tr.X - vp.bl.X) * (vp.tr.Y - vp.bl.Y)) / mPolyArea
        If areafrac > DrawAllFraction Then areafrac = 1
        
        ' check that the viewport touchs the Root Cell
        If vp.tr.X < mbl.X Then
        ElseIf vp.bl.X > mtr.X Then
        ElseIf vp.tr.Y < mbl.Y Then
        ElseIf vp.bl.Y > mtr.Y Then
        Else
            ' fast way to get all references if views entrie object
            If vp.bl.X <= mbl.X Then
               If vp.tr.X >= mtr.X Then
                  If vp.bl.Y <= mbl.Y Then
                     If vp.tr.Y >= mtr.Y Then
                        ' The Viewport covers the whole of the Root Cell
                        ' simply output ALL the references 0 to maxReference
                        ReDim QuadOutput(maxReference)
                        For i = 0 To maxReference
                            QuadOutput(i) = i
                        Next i
                        nQuadOutput = maxReference
                        Viewport_Top = True   ' causes viewport to exit sub
                        Exit Function
                     End If
                  End If
               End If
            End If
            
            ' so we know the viewport and Root Cell touch...
            ' if areafrac = 1 (ie vp size most of poly)
            If areafrac = 1 Then
                ' find the area of intersection
                ' The maths here is simple, the concept / method is not.
                
                ' get the highest (right most) left hand edge
                If mbl.X > vp.bl.X Then
                    X(0) = mbl.X
                Else
                    X(0) = vp.bl.X
                End If
                ' get the lowest (left most) right hand edge
                If mtr.X < vp.tr.X Then
                    X(1) = mtr.X
                Else
                    X(1) = vp.tr.X
                End If
                ' get the highest bottom  edge
                If mbl.Y > vp.bl.Y Then
                    Y(0) = mbl.Y
                Else
                    Y(0) = vp.bl.Y
                End If
                ' get the lowest top edge
                If mtr.Y < vp.tr.Y Then
                    Y(1) = mtr.Y
                Else
                    Y(1) = vp.tr.Y
                End If
                
                ' the amount of intersection in X is then given by (x(1) - x(0))
                ' the amount of intersection in Y is then given by (y(1) - y(0))
                ' Calculate the Area of the Root Cell that IS (not COULD BE) covered by the viewport
                intarea = (X(1) - X(0)) * (Y(1) - Y(0)) / mPolyArea
                If intarea > DrawAllFraction Then ' is the covered area greater than our
                    ReDim QuadOutput(maxReference)
                    For i = 0 To maxReference
                        QuadOutput(i) = i
                    Next i
                    nQuadOutput = maxReference
                    Viewport_Top = True   ' causes viewport to exit sub
                    Exit Function
                End If
            End If
          End If
        
        InitRedim = maxReference * areafrac
        If InitRedim <= (noSubDivide * 2) Then InitRedim = (noSubDivide * 2)
        AddRedim = InitRedim / 2
        ReDim QuadOutput(InitRedim)
        nQuadOutput = 0
        ubQuadOutput = InitRedim
        ReDim blnRef(maxReference)

End Function

