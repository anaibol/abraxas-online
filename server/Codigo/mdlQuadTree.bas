Attribute VB_Name = "mdlQuadTree"
Option Explicit

Public Type coord
    X As Double
    Y As Double
End Type

Public Type Rectangle
    bl As coord             ' Bottom Left
    tr As coord             ' Top Right
End Type

'"Constants"
Const DrawAllPercentage As Double = 95   ' Used in Viewport
' If more than this percentage of the first cell is covered by the viewport
' it will simply return ALL the objects. Attempting to reduce "worst case"
Public Const DrawAllFraction = DrawAllPercentage / 100

'Inputs
Public PolyRects(0 To 1000 - 1) As Rectangle     ' array of rectangles (OBJECTS)
Public nPolyRects As Long
Public plyrefs(0 To 1000 - 1) As Long            ' Array to hold initial references (usually 0 to nObjects)

Public vp As Rectangle              ' Viewport Rectangle
Public pt As coord                  ' Point to Test

' Outputs
Public QuadOutput() As Long   ' The Output Array
Public nQuadOutput As Long    ' Number of references in the output array
Public ubQuadOutput As Long   ' uBound of the Output array (much quicker than Ubound())

' Other
Public blnRef() As Integer    ' Array for eliminating repeats
Public InitRedim As Long      ' Initial Size of the Output Array
Public AddRedim As Long       ' Number of Extra Elements to add when needed
Public maxReference As Long   ' The total number of objects
Public noSubDivide As Long    ' Maximum number of objects in a "lowest" cell
                              ' Currently Calculated by CreateTree

' A Class Instance
Public tstQuad As clsQuadTree

' I included my Sub for filling the PolyRects() and plyrefs() array in here.
Public Sub PopulatePolyRects()

  Set tstQuad = New clsQuadTree
  tstQuad.IsTop = True
  'tstQuad.CreateTree plyrefs(), MAX_OBJECTS - 1, 0#, 0#, 100000#, 100000#
  
  
  'Picture1.Scale (0, 100000)-(100000, 0)
  'scaleSize = Picture1.ScaleWidth
  'InitialScaleSize = scaleSize / Screen.TwipsPerPixelX
  
  'DrawObjectsInViewPort

' Sample Code

' To get started
   'Set tstQuad = New clsQuadTree
   'tstQuad.IsTop = True
   'tstQuad.CreateTree plyrefs(), nPolyRects - 1, MinXBorder, MinYBorder, MaxXBorder, MaxYBorder

' To Finish
   'Set tstQuad = Nothing


End Sub
