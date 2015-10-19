Attribute VB_Name = "modDeshacer"
Option Explicit

' Deshacer
Public Const maxDeshacer As Integer = 10
Public MapData_Deshacer(1 To maxDeshacer, 1 To 100, 1 To 100) As MapBlock
Public MapData_Deshacer_Info(1 To maxDeshacer) As Boolean

Public Sub Deshacer_Clear()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Dim i As Integer
' Vacio todos los campos afectados
For i = 1 To maxDeshacer
    MapData_Deshacer_Info(i) = True
Next
' no ahi que deshacer
frmMain.mnuDeshacer.Enabled = False
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByVal Desc As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Exit Sub
Dim i As Integer
Dim F As Integer
Dim j As Integer
' Desplazo todos los deshacer uno hacia atras
For i = maxDeshacer To 2 Step -1
    For F = 1 To MapInfo.dX
        For j = 1 To MapInfo.dY
            MapData_Deshacer(i, F, j) = MapData_Deshacer(i - 1, F, j)
        Next
    Next
    MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i - 1)
Next
' Guardo los valores
For F = 1 To MapInfo.dX
    For j = 1 To MapInfo.dY
        MapData_Deshacer(1, F, j) = MapData(F, j)
    Next
Next
MapData_Deshacer_Info(1) = False
frmMain.mnuDeshacer.Caption = "&Deshacer"
frmMain.mnuDeshacer.Enabled = True
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************

Exit Sub
Dim i As Integer
Dim F As Integer
Dim j As Integer
Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
Dim npci As Integer
If MapData_Deshacer_Info(1) = False Then
    ' Aplico deshacer
    For F = 1 To MapInfo.dX
        For j = 1 To MapInfo.dY
            If (MapData(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex <> MapData_Deshacer(1, F, j).NPCIndex) Or (MapData(F, j).NPCIndex <> 0 And MapData_Deshacer(1, F, j).NPCIndex = 0) Then
                ' Si ahi un NPC, y en el deshacer es otro lo borramos
                ' (o) Si aun no NPC y en el deshacer no esta
                MapData(F, j).NPCIndex = 0
                Call Engine.Char_Erase(MapData(F, j).CharIndex)
            End If
            If MapData_Deshacer(1, F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex = 0 Then
                ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                
                npci = MapData_Deshacer(1, F, j).NPCIndex
                MapData(F, j) = MapData_Deshacer(1, F, j)
                
                Body = NpcData(npci).Body
                Head = NpcData(npci).Head
                Heading = NpcData(npci).Heading
                Call Engine.Char_Make(NextOpenChar(), Body, Head, Heading, F, j, 2, 2, 2)
            Else
                MapData(F, j) = MapData_Deshacer(1, F, j)
            End If
        Next
    Next
    MapData_Deshacer_Info(1) = True
    ' Desplazo todos los deshacer uno hacia adelante
    For i = 1 To maxDeshacer - 1
        For F = 1 To MapInfo.dX
            For j = 1 To MapInfo.dY
                MapData_Deshacer(i, F, j) = MapData_Deshacer(i + 1, F, j)
            Next
        Next
        MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
    Next
    ' borro el ultimo
    MapData_Deshacer_Info(maxDeshacer) = True
    ' ahi para deshacer?
    If MapData_Deshacer_Info(1) = True Then
        frmMain.mnuDeshacer.Caption = "&Deshacer (no ahi nada que deshacer)"
        frmMain.mnuDeshacer.Enabled = False
    Else
        frmMain.mnuDeshacer.Caption = "&Deshacer"
        frmMain.mnuDeshacer.Enabled = True
    End If
Else
    MsgBox "No ahi acciones para deshacer", vbInformation
End If
End Sub

