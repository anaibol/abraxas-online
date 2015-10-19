VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'clsCustomMessages.cls - User-defined messages with shortcut.
'Allows the creation and serialization of user defined messages to be used with shortcuts.

Option Explicit
Option Base 0

'Defines the number of messages that can be created.
Private Const MESSAGE_COUNT As Byte = 10

'Defines the maximum length allowed for any custom message.
Private Const MAX_MESSAGE_LENGTH As Byte = 120

'The Error code thrown when the message being set is invalid.
Private Const INVALID_MESSAGE As Long = vbObjectError + 123

'Relative path from the application in which custom message are stoRed.
Private Const MESSAGE_FILE_PATH As String = "/Data/messages.txt"

'Array of defined messages.
Private CustomMessages(MESSAGE_COUNT - 1) As String

'Class constructor. Initializes a new clsCustomMessage

Private Sub Class_Initialize()
'Initializes the messages by loading them from the config file

On Error GoTo ErrHandler
    Dim handle As Integer
    Dim i As Long
    
    'Load custom messages
    handle = FreeFile()
    
    Open App.path & MESSAGE_FILE_PATH For Input As handle
    
    For i = 0 To MESSAGE_COUNT - 1
        Line Input #handle, CustomMessages(i)
        
        'Purge invalid messages
        If Not MessageIsValid(CustomMessages(i)) Then
            CustomMessages(i) = vbNullString
        End If
    Next i
    
    Close handle
    
    'If the file does not exist (only possible Error),
    'then we have nothing to worry. All messages will be empty
ErrHandler:
End Sub

Private Sub Class_Terminate()
'Class desconstructor. Serializes messages to
'Stores all messages back into config file for later use.

    Dim handle As Integer
    Dim i As Long
    Dim directories() As String
    Dim path As String
    
    'Make sure the path exists
    directories = Split(MESSAGE_FILE_PATH, "/")
    path = App.path
    
    For i = 0 To UBound(directories()) - 1
        path = path & "/" & directories(i)
        
        If LenB(Dir(path, vbDirectory)) = 0 Then
            Call MkDir(path)
        End If
    Next i
    
    'Serialize custom messages
    handle = FreeFile()
    
    Open App.path & MESSAGE_FILE_PATH For Output As handle
    
    For i = 0 To MESSAGE_COUNT - 1
        Print #handle, CustomMessages(i)
    Next i
    
    Close handle
End Sub

Public Property Get Message(ByVal Index As Integer) As String
'Retrieves the requested message.
    If Index < 0 Or Index >= MESSAGE_COUNT Then
        Exit Property
    End If
   
    Message = CustomMessages(Index)
End Property

Public Property Let Message(ByVal Index As Integer, ByVal nValue As String)
'Sets the given message index with the requested message.

    If Index < 0 Or Index >= MESSAGE_COUNT Then
        Exit Property
    End If
    
    If Not MessageIsValid(nValue) Then
        Call Err.Raise(INVALID_MESSAGE)
        Exit Property
    End If
    
    CustomMessages(Index) = nValue
End Property

Public Property Get InvalidMessageErrCode() As Long
'Retrieves the INVALID_MESSAGE Error code.

    InvalidMessageErrCode = INVALID_MESSAGE
End Property

Private Function MessageIsValid(ByRef msg As String) As Boolean
'Erased line "msg=StrConv(msg,vbUnicode)", because
'it was causing problems and detected as many messages
'as wrong with nothing bad on them

    Dim i As Long
    Dim Char As String * 1
    Dim ascii As Byte
    
    'Get rid of spaces
    msg = Trim$(msg)
        
    'Empty messages are valid
    If LenB(msg) = 0 Then
        MessageIsValid = True
        Exit Function
    End If
    
    For i = 1 To Len(msg)
        Char = mid$(msg, i, 1)
        ascii = Asc(Char)
        
        'check for invalid chars...
        If ascii < 31 Then
           Exit Function
        End If
        
        If ascii > 126 And Char <> "á" And Char <> "Á" And Char <> "ä" And Char <> "Ä" _
          And Char <> "â" And Char <> "Â" And Char <> "à" And Char <> "À" And Char <> "é" _
          And Char <> "É" And Char <> "ë" And Char <> "Ë" And Char <> "ê" And Char <> "Ê" _
          And Char <> "è" And Char <> "È" And Char <> "í" And Char <> "Í" And Char <> "ï" _
          And Char <> "Ï" And Char <> "î" And Char <> "Î" And Char <> "ì" And Char <> "Ì" _
          And Char <> "ó" And Char <> "Ó" And Char <> "ö" And Char <> "Ö" And Char <> "ô" _
          And Char <> "Ô" And Char <> "ò" And Char <> "Ò" And Char <> "ú" And Char <> "Ú" _
          And Char <> "ü" And Char <> "Ü" And Char <> "û" And Char <> "Û" And Char <> "ù" _
          And Char <> "Ù" And Char <> "ñ" And Char <> "Ñ" And Char <> "¡" And Char <> "¿" Then
            Exit Function
        End If
    Next i
    
    'Discard extremely long messages
    If Len(msg) > MAX_MESSAGE_LENGTH Then
        Exit Function
    End If
    
    MessageIsValid = True
End Function
