VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'The MainTimer of the game.
'Controls client´s intervals.

Option Explicit

'How many timers we are going to use-
Private Const CantidadTIMERS As Byte = 14

Public Enum TimersIndex
    Attack = 1
    Work = 2
    UseItemWithU = 3
    UseItemWithDblClick = 4
    SendRPU = 5
    CastSpell = 6
    Arrows = 7
    CastAttack = 8
    Drop = 9
    BuySell = 10
    PublicMessage = 11
    Talk = 12
    Medit = 13
    RandomName = 14
End Enum

Private Type Timer
    Interval As Long
    CurrentTick As Long
    StartTick As Long
    Run As Boolean
End Type

'Timers
Dim Timer(1 To CantidadTIMERS) As Timer



Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Sub
    End If
    Timer(TimerIndex).Interval = Interval
End Sub

Public Function GetInterval(ByVal TimerIndex As TimersIndex) As Long
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Function
    End If
    GetInterval = Timer(TimerIndex).Interval
End Function

Public Sub Start(ByVal TimerIndex As TimersIndex)
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Sub
    End If
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
End Sub

Public Sub Continue(ByVal TimerIndex As TimersIndex)
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Sub
    End If
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
    
    'We´ll set the StartTick again
    Timer(TimerIndex).StartTick = Timer(TimerIndex).StartTick + Timer(TimerIndex).CurrentTick
End Sub

Public Sub Pause(ByVal TimerIndex As TimersIndex)
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Sub
    End If
    
    Timer(TimerIndex).Run = False 'Don´t allow the timer to run.
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calculates acutual tick
End Sub

Public Function Check(ByVal TimerIndex As TimersIndex, Optional Restart As Boolean = True) As Boolean
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Function
    End If
    
    If Not Timer(TimerIndex).Run Then
        Exit Function 'If we weren´t allowed to run, exit function
    End If
    
    Timer(TimerIndex).CurrentTick = GetTickCount - Timer(TimerIndex).StartTick 'Calcutates CurrentTick
    
    If Timer(TimerIndex).CurrentTick >= Timer(TimerIndex).Interval Then
        Check = True 'We have Ticked!
        If Restart Then
            Timer(TimerIndex).StartTick = GetTickCount 'Restart Timer (Nicer than calling Restart() )
            If (TimerIndex = TimersIndex.Attack) Or (TimerIndex = TimersIndex.CastSpell) Then
                Timer(TimersIndex.CastAttack).StartTick = GetTickCount 'Set Cast-Attack interval
            ElseIf TimerIndex = TimersIndex.CastAttack Then
                Timer(TimersIndex.Attack).StartTick = GetTickCount 'Restart Attack interval
                Timer(TimersIndex.CastSpell).StartTick = GetTickCount 'Restart Magic interval
            End If
        End If
    End If
End Function

Public Sub Restart(ByVal TimerIndex As TimersIndex)
    If TimerIndex < 1 Or TimerIndex > CantidadTIMERS Then
        Exit Sub
    End If
    Timer(TimerIndex).StartTick = GetTickCount
End Sub
