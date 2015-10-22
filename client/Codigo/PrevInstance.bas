Attribute VB_Name = "modPrevInstance"
'Prevents multiple instances of the game running on the same computer.

Option Explicit

'Declaration of the Win32 API Public function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const Error_ALREADY_EXISTS = 183&

Private mutexHID As Long
Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'Creates a Named Mutex. private function, since we will use it just to check if a previous instance of the app is running.
    
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = Error_ALREADY_EXISTS) 'check if the mutex already existed
End Function

Public Function FindPreviousInstance() As Boolean
'Checks if there's another instance of the app running, returns True if there is or False otherwise.

    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("aguantemegadeth") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
    End If
End Function

Public Sub ReleaseInstance()
'Closes the client, allowing other instances to be open.

    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub
