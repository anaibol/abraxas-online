VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Note - this is not the declare in the API viewer - modify lplpVoid to be
'Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 'color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmplanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Const CF_BITMAP = 2

'Handle to the current DIBSection:
Private m_hDIb As Long
'Handle to the old bitmap in the DC, for clear up:
Private m_hbmpOld As Long
'Handle to the Device context holding the DIBSection:
Private m_hDC As Long
'Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
'Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib > 0)
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture)
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhbmpOld As Long
Dim tbmp As BITMAP
Dim lhWnd As Long
    
    GetObjectAPI picThis.handle, Len(tbmp), tbmp
    If (Create(tbmp.bmWidth, tbmp.bmHeight)) Then
        lhWnd = GetDesktopWindow()
        lhDCDesktop = GetDC(lhWnd)
        If (lhDCDesktop > 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            ReleaseDC lhWnd, lhDCDesktop
            If (lhDC > 0) Then
                lhbmpOld = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhbmpOld
                DeleteDC lhDC
            End If
        End If
    End If
End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC > 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hbmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    'Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lhDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then
        lSrcWidth = m_tBI.bmiHeader.biWidth
    End If
    
    If lSrcHeight < 0 Then
        lSrcHeight = m_tBI.bmiHeader.biHeight
    End If
    
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Sub PaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then
        lDestWidth = m_tBI.bmiHeader.biWidth
    End If
    
    If (lDestHeight < 0) Then
        lDestHeight = m_tBI.bmiHeader.biHeight
    End If
    
    BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property

Public Sub ClearUp()
    If (m_hDC > 0) Then
        If (m_hDIb > 0) Then
            SelectObject m_hDC, m_hbmpOld
            DeleteObject m_hDIb
        End If
        DeleteDC m_hDC
    End If
    m_hDC = 0: m_hDIb = 0: m_hbmpOld = 0: m_lPtr = 0
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub




