Attribute VB_Name = "modRessources"
Option Explicit

'Private hOldCursor As Long

'Private declare function CopyCursor Lib "user32" Alias "CopyIcon" (ByVal hcur As Long) As Long
'Private declare function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'Private declare function GetCursor Lib "user32" () As Long

Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Private Declare Function LoadCursorFromFile Lib "user32" Alias _
    "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
    
Private Declare Function GetTempFilename Lib "kernel32" _
    Alias "GetTempFileNameA" ( _
    ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFilename As String _
    ) As Long

Private Declare Function GetTempPath Lib "kernel32" _
    Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String _
    ) As Long
    
Dim Font1 As String
Dim Font2 As String

Public Sub LoadRessources()

    Dim sFileName As String
    
    If GetTempFile(vbNullString, "~rs", 0, sFileName) Then
        If Not SaveResItemToDisk(101, "Custom", sFileName) Then
        
            'hOldCursor = CopyCursor(GetCursor())
            
            Call SetSystemCursor(LoadCursorFromFile(sFileName), 32512)
            
            'Delete the temp file
            Kill sFileName
        End If
    End If
        
    If GetTempFile(vbNullString, "~rs", 0, Font1) Then
    
        If Not SaveResItemToDisk(102, "Custom", Font1) Then
        
            Call AddFontResource(Font1)
            
            'Delete the temp file
            'Kill Font1
        End If
    End If
    
    If GetTempFile(vbNullString, "~rs", 0, Font2) Then
        If Not SaveResItemToDisk(103, "Custom", Font2) Then
        
            Call AddFontResource(Font2)

            'Delete the temp file
            'Kill Font2
        End If
    End If
    
    Call SendMessage(&HFFFF&, &H1D, 0, 0)

End Sub

Public Sub UnloadRessources()
    'Dim Ob As Object
    
    'Set Ob = CreateObject("Wscript.Shell")

    'SetSystemCursor LoadCursorFromFile(Replace(Ob.RegRead("HKEY_CURRENT_USER\Control Panel\Cursors\Arrow"), "%SystemRoot%", Ob.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRoot"))), 32512

    'SetSystemCursor hOldCursor, 32512
        
    Call SystemParametersInfo(&H57, 0, ByVal 0&, &H2)
        
    Exit Sub
    
    Call RemoveFontResource(Font1)
    Call RemoveFontResource(Font2)
    
    Call SendMessage(&HFFFF&, &H1D, 0, 0)
End Sub

Public Function GetTempFile( _
    ByVal strDestPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Integer, _
    lpTempFilename As String _
    ) As Boolean
    '==========================================================================
    'Get a temporary filename for a specified drive and filename prefix
    'PARAMETERS:
    'strDestPath - Location where temporary file will be created.  If this
    'is an empty string, then the location specified by the
    'tmp or temp environment variable is used.
    'lpPrefixString - First three Chars of this string will be part of
    'temporary file name returned.
    'wUnique - Set to 0 to create unique filename.  Can also set to integer,
    'in which case temp file name is returned with that integer
    'as part of the name.
    'lpTempFilename - Temporary file name is returned as this variable.
    'RETURN:
    'True if Public function succeeds; false otherwise
    '==========================================================================
    
    If strDestPath = vbNullString Then
        'No destination was specified, use the temp directory.
        strDestPath = String(255, vbNullChar)
        If GetTempPath(255, strDestPath) = 0 Then
            GetTempFile = False
            Exit Function
        End If
    End If
    lpTempFilename = String(255, vbNullChar)
    GetTempFile = GetTempFilename(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function

Public Function SaveResItemToDisk( _
            ByVal iResourceNum As Integer, _
            ByVal sResourceObjType As String, _
            ByVal sDestFileName As String _
            ) As Long
    '=============================================
    'Saves a resource Item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    'Example Call:
    'iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")
    
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    
    On Error GoTo SaveResItemToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceObjType)
    
    'Get Free File Handle
    iFileNumOut = FreeFile
    
    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut
        
        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData
    
    'Close the file
    Close #iFileNumOut
    
    'Return 0 for success
    SaveResItemToDisk = 0
    
    Exit Function
SaveResItemToDisk_err:
    'Return error number
    SaveResItemToDisk = Err.Number
End Function

Public Function StripTerminator(ByVal strString As String) As String
    '==========================================================
    'Returns a string without any zero terminator.  Typically,
    'this was a string returned by a Windows API call.
    '
    'IN: [strString] - String to remove terminator from
    '
    'Returns: The value of the string passed in minus any
    'terminating zero.
    '==========================================================
    
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


