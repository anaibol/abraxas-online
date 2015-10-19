VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Customizes User Keys.

Option Explicit

'
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 26

'
'Error Raised in case the key file doesn't exist
Private Const Error_NO_KEY_FILE As Long = vbObjectError + 16

'
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


'
'Index of keys on array
'
'@param mKeyUp Key for moving Char up
'@param mKeyDown Key for moving Char down
'@param mKeyLeft Key for moving Char left
'@param mKeyRight Key for moving Char right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling Chars'names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users'gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the Char
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyToggleFPS Key for toggling the FPS counter
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyToggleFPS
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
End Enum

Public Function ReadableName(ByVal KeyCode As Integer) As String
'Shows a readable name for a specific Key code

    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado numérico)"
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
            ReadableName = "Entrar (teclado numérico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Ctrl"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Entrar"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
        
        Case Else
            'In all other cases, key is consideRed not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

Public Sub LoadCustomKeys()

    Dim i As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    'Inits keyboard config with default configurations
    Call LoadDefaults
        
    Open DataPath & "BindKeys.bin" For Binary Access Read Lock Read Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Get FileNum, , mappedKeys(i)
        If LenB(ReadableName(mappedKeys(i))) = 0 Then
            Close FileNum
            Call Err.Raise(Error_NO_KEY_FILE)
            Exit Sub
        End If
    Next i
    
    Close FileNum
End Sub

Public Sub SaveCustomKeys()
'Saves custom keys to a file

    Dim FileNum As Integer
    Dim i As Long
    FileNum = FreeFile()
    
    Open DataPath & "BindKeys.bin" For Binary Access Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Put FileNum, , mappedKeys(i)
    Next i
    
    Close #FileNum
End Sub

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'Gets a key code
   
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then
        Exit Property
    End If
    
    BindedKey = mappedKeys(Index)
End Property

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'Sets a key code
    
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then
        Exit Property
    End If
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(Index) = NewVal
End Property

Public Sub LoadDefaults()
'Loads default keys
    
    mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
    mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
    mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyMultiply
    mappedKeys(eKeyType.mKeyToggleResuscitationSafe) = KeyCodeConstants.vbKeyEnd
    mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
    mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
    mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
    mappedKeys(eKeyType.mKeyToggleFPS) = KeyCodeConstants.vbKeyF4
    mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF2
    mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
    mappedKeys(eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
    mappedKeys(eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
    mappedKeys(eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
    mappedKeys(eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyF12
End Sub

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'Returns whether a key is binded to an action
    
    Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

Public Property Get Count() As Byte
'Retrieves the number of customizable keys

    Count = NUM_CUSTOM_KEYS
End Property

Private Sub Class_Initialize()
'Constructor. Loads Keys from file

On Error GoTo ErrHand:
    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.Number = Error_NO_KEY_FILE Then
        Call LoadDefaults
        Call SaveCustomKeys
        Resume Next
    End If
End Sub

Private Sub Class_Terminate()
'Destructor. Saves Keys to file

    Call SaveCustomKeys
End Sub
