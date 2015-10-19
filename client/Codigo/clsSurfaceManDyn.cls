VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As DirectDrawSurface7
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private DirectDraw As DirectDraw7

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function SetDIBitsToDevice Lib "gdi32" _
    (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, _
    ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal uStartScan As Long, ByVal CScanLine As Long, _
    lpBits As Any, lpbmi As BITMAPINFO, ByVal fuColorUse As Long) As Long

Private Sub Class_Initialize()
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
End Sub

Private Sub Class_Terminate()
'Clean up

    Dim i  As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
            Next j
    
    'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub


Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'Initializes the manager

    Set DirectDraw = DD
    
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   '16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       '4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As DirectDrawSurface7
'Retrieves the requested texture
    
    Dim i As Long
    
    'Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Property
    End If
        Next i
    End With
    
    'Not in memory, load it!
    Set clsSurfaceManager_Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As DirectDrawSurface7
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position

On Error Resume Next
'On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim ddsd As DDSURFACEDESC2
    Dim ddck As DDCOLORKEY
    Dim data() As Byte
    Dim bmpInfo As BITMAPINFO
    Dim sDC As Long
    
    'Get Bitmap
    'Call Get_Bitmap(ResourcePath, CStr(fileIndex) & ".bmp", bmpInfo, data)

    'Set up the surface desc
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    ddsd.lWidth = bmpInfo.bmiHeader.biWidth
    ddsd.lHeight = bmpInfo.bmiHeader.biHeight
        
    With newSurface
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        
        'Load surface
        'Set .Surface = DirectDraw.CreateSurface(ddsd)
        
        Set .Surface = DirectDraw.CreateSurfaceFromFile(GrhPath & CStr(fileIndex) & ".bmp", ddsd)
        
        sDC = .Surface.GetDC
        
'        Call SetDIBitsToDevice(sDC, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, 0, 0, 0, bmpInfo.bmiHeader.biHeight, data(0), bmpInfo, DIB_RGB_COLORS)
        
        Call .Surface.ReleaseDC(sDC)
        
        'Set colorkey
        ddck.high = 0
        ddck.low = 0
        Call .Surface.SetColorKey(DDCKEY_SRCBLT, ddck)
        
        'Retrieve the updated surface desc
        Call .Surface.GetSurfaceDesc(ddsd)
    End With
    
    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
    
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + ddsd.lHeight * ddsd.lPitch
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
    Exit Function

ErrHandler:
End Function

Private Function RemoveLRU() As Boolean
'Removes the Least Recently Used surface to make some room for new ones

    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim ddsd As DDSURFACEDESC2
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).lastAccess Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).lastAccess
                End If
            Next j
        End With
    Next i
    
    If LRUj Then
        RemoveLRU = True
        'Retrieve the surface desc
        Call surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetSurfaceDesc(ddsd)
        
        'Remove it
        Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
        surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
        
        'Move back the list (if necessary)
        With surfaceList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
        
        'Update the used bytes
        usedBytes = usedBytes - ddsd.lHeight * ddsd.lPitch
    End If
End Function
