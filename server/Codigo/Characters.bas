Attribute VB_Name = "Characters"
Option Explicit

'Value representing invalid Indexes.
Public Const INVALID_Index As Integer = 0

Public Function CharIndexToUserIndex(ByVal CharIndex As Integer) As Integer
'Takes a CharIndex and transforms it into a UserIndex. Returns INVALID_Index in case of error.
    CharIndexToUserIndex = CharList(CharIndex)
    
    If CharIndexToUserIndex < 1 Or CharIndexToUserIndex > MaxPoblacion Then
        CharIndexToUserIndex = INVALID_Index
        Exit Function
    End If
    
    If UserList(CharIndexToUserIndex).Char.CharIndex <> CharIndex Then
        CharIndexToUserIndex = INVALID_Index
        Exit Function
    End If
End Function
