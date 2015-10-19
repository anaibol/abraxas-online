VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'clsDialogs.cls - Chats over head manager.
'Creates and manipulates Chats over heads, rendering and updating them.

Option Explicit
Option Base 0

'Maximum number of dialogs that can exist.
Private Const MAX_DIALOGS As Byte = 100

'Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH As Byte = 18

'Number of milliseconds to add to the lifetime per dialog Char
Private Const MS_PER_CHAR As Byte = 100

'Number of extra milliseconds to add to the lifetime of a new dialog
Private Const MS_ADD_EXTRA As Integer = 5000

Private Type dialog
    textLine()  As String
    X           As Integer
    Y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    color       As Long
    renderable  As Boolean
    SUBe As Integer 'Efecto (INTEGER!)
End Type

'Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1) As dialog

'The number of dialogs being used at the moment.
Private dialogCount As Byte

'The size of the font being used to render
Private usedFont As StdFont

Private Sub Class_Initialize()
'Initializes the dialog instance
    dialogCount = 0
End Sub

Private Sub Class_Terminate()
'Clean up
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
'Returns the index of the dialog in the list, or the negation
'of the position were it should be if not found (for binary insertion)

    Dim Min As Long
    Dim max As Long
    Dim mid As Long
    
    Min = 0
    max = dialogCount - 1
    
    Do While Min <= max
        mid = (Min + max) * 0.5
        
        If dialogs(mid).CharIndex < CharIndex Then
            Min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not Min
End Function

Public Sub CreateDialog(ByRef Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'Creates a new dialog. If the user already had a dialog, then it's replaced,
'otherwise a new Slot is used.
    
    Dim Slot As Integer
    Dim i As Long
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then
        If dialogCount = MAX_DIALOGS Then
            Exit Sub  'Out of space! Should never happen....
        End If
        
        'We need to add it. Get insertion index and move list backwards.
        Slot = Not Slot
        
        For i = dialogCount To Slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1
    End If
    
    With dialogs(Slot)
        .CharIndex = CharIndex
        .color = color
        .renderable = False
        .startTime = GetTickCount()
        .textLine = FormatChat(Chat)
        .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(Chat))
        .SUBe = 8
    End With
    
    If CharIndex = UserCharIndex Then
        Call RemoveExp
        Call RemoveDamage
        Call RemoveGld
    End If
End Sub

Private Function FormatChat(ByRef Chat As String) As String()
'Formats a dialog into different text lines.

    Dim word As String
    Dim curPos As Long
    Dim length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim ChatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    length = Len(Chat)
    acumLength = 0
    lineLength = -1
    ReDim ChatLines(FieldCount(Chat, 32)) As String
    
    'Start formating
    Do While acumLength < length
        word = ReadField(curPos, Chat, 32)
        
        wordLength = Len(word)
        
        'Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            ChatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            'Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                'Put it in the next line
                curLine = curLine + 1
                ChatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                ChatLines(curLine) = ChatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    'If it's only one line, center text
    If curLine = 0 And length < MAX_LENGTH Then
        ChatLines(curLine) = String((MAX_LENGTH - length) * 0.5 + 1, " ") & ChatLines(curLine)
    End If
    
    'Resize array to fit
    ReDim Preserve ChatLines(curLine) As String
    
    FormatChat = ChatLines
End Function

Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal CharIndex As Integer)
'Updates a dialog's position.

    Dim Slot As Integer
    
    Slot = BinarySearch(CharIndex)
    
    If Slot >= 0 Then
        With dialogs(Slot)
            .X = X - (MAX_LENGTH * 0.5) * 3
            .Y = Y - (UBound(.textLine()) * 3)
            .renderable = True
        End With
    End If
End Sub

Public Sub Render()
'Displays all renderable dialogs on screen.

    Dim i As Long
    Dim j As Long
    Dim offset As Integer
    Dim removedDialogs As Long
    
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)
            If (GetTickCount() - .startTime) >= .lifeTime Then
                Call RemoveDialog(.CharIndex)
                
                removedDialogs = removedDialogs + 1
            ElseIf .renderable Then
                If .lifeTime >= 292 Then
                    .SUBe = .SUBe - 1
                    If .SUBe > 0 Then
                        .Y = .Y + .SUBe
                    End If
                End If
                offset = -(usedFont.size + 2) * UBound(.textLine())
                
                'Render it
                For j = 0 To UBound(.textLine())
                    Call RenderText(.X, .Y + offset - 2, .textLine(j), .color)
                    offset = offset + usedFont.size + 5
                Next j
                
                'Won't render it unless it's updated in next frmae
                .renderable = False
            End If
        End With
    Next i
End Sub

Public Sub RemoveDialog(ByVal CharIndex As Integer)
'Removes the dialog from the requested char.

    Dim Slot As Integer
    Dim i As Long
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then
        Exit Sub
    End If
    
    For i = Slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
End Sub

Public Sub RemoveAllDialogs()
'Removes all dialogs from all chars.
    dialogCount = 0
End Sub

Public Property Get font() As StdFont
'Retrieves the font being used to render dialogs.
    font = usedFont
End Property

Public Property Let font(ByRef nValue As StdFont)
'Sets the font being used to render dialogs.
    Set usedFont = nValue
End Property
