Attribute VB_Name = "Statistics"
Option Explicit

Private Type trainningData
    startTick As Long
    trainningTime As Long
End Type

Private Type fragLvlRace
    matrix(1 To 50, 1 To 5) As Long
End Type

Private Type fragLvlLvl
    matrix(1 To 50, 1 To 50) As Long
End Type

Private trainningInfo() As trainningData

Private fragLvlRaceData(1 To 7) As fragLvlRace
Private fragLvlLvlData(1 To 7) As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long

'Currency just in case.... chats are way TOO often...
Private keyOcurrencies(255) As Currency

Public Sub Initialize()
    ReDim trainningInfo(1 To MaxPoblacion) As trainningData
End Sub

Public Sub UserConnected(ByVal UserIndex As Integer)
    'A new user connected, load it's trainning time count
    trainningInfo(UserIndex).trainningTime = Val(GetVar(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "RESEARCH", "TrainningTime", 30))
    
    trainningInfo(UserIndex).startTick = (GetTickCount() And &H7FFFFFFF)
End Sub

Public Sub UserDisconnected(ByVal UserIndex As Integer)
    With trainningInfo(UserIndex)
        'Update trainning time
        .trainningTime = .trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) \ 1000
        
        .startTick = (GetTickCount() And &H7FFFFFFF)
        
        'Store info in char file
        Call WriteVar(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", "RESEARCH", "TrainningTime", CStr(.trainningTime))
    End With
End Sub

Public Sub UserLevelUp(ByVal UserIndex As Integer)
    Dim handle As Integer
    handle = FreeFile()
    
    With trainningInfo(UserIndex)
        'Log the data
        Open App.Path & "/logs/statistics.log" For Append Shared As handle
        
        Print #handle, UCase$(UserList(UserIndex).Name) & " completó el nivel " & CStr(UserList(UserIndex).Stats.Elv) & " en " & CStr(.trainningTime + ((GetTickCount() And &H7FFFFFFF) - .startTick) \ 1000) & " segundos."
        
        Close handle
        
        'Reset data
        .trainningTime = 0
        .startTick = (GetTickCount() And &H7FFFFFFF)
    End With
End Sub
