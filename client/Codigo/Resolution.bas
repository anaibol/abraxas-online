Attribute VB_Name = "modResolution"
'Resolution.bas - Performs resolution changes.

Option Explicit

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPixel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private oldDepth As Integer
Private oldFrequency As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

'TODO : Change this to not depend on any external public variable using args instead!

Public Sub SetResolution()
'Changes the display resolution if needed.

    Dim lRes As Long
    Dim MidevM As typDevMODE
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
                    
    With MidevM
        If ChangeResolution Then
            oldDepth = .dmBitsPerPixel
            oldFrequency = .dmDisplayFrequency
    
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
        Else
            .dmFields = DM_BITSPERPEL
            frmMain.WindowState = vbNormal
        End If
        
        If AlphaBActivated Then
            .dmBitsPerPixel = 16
        End If
    End With
        
    lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
End Sub

Public Sub ResetResolution()
'Changes the display resolution if needed.

    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
    With typDevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
        .dmPelsWidth = 0
        .dmPelsHeight = 0
        .dmBitsPerPixel = oldDepth
        .dmDisplayFrequency = oldFrequency
    End With
        
    lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
End Sub

Public Function ResolucionActual() As Boolean

    Dim lRes As Long
    Dim MidevM As typDevMODE
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
                    
    With MidevM
        oldDepth = .dmBitsPerPixel
        oldFrequency = .dmDisplayFrequency

        If .dmPelsWidth = 800 And _
        .dmPelsHeight = 600 Then
            ResolucionActual = 1
        Else
            ResolucionActual = 0
        End If
    End With

End Function
