Attribute VB_Name = "modAutoInsert"
Option Explicit

'Basicos
Private Const CAM_VERT_S As Byte = 35
Private Const CAM_VERT_UP As Byte = 34
Private Const CAM_VERT_DN As Byte = 36

Private Const CAM_HORI_S As Byte = 32
Private Const CAM_HORI_LF As Byte = 31
Private Const CAM_HORI_RG As Byte = 33

'Curvas
Private Const CAM_UP_LEFT As Byte = 40
Private Const CAM_UP_RIGHT As Byte = 39
Private Const CAM_DN_LEFT As Byte = 38
Private Const CAM_DN_RIGHT As Byte = 37


Private Const CAM_TOTAL As Byte = 45


'Public Sub Insert_Cam()
'    Dim loopC As Boolean
'
'    Dim X As Integer, Y As Integer
'
'    For Y = 1 To 100
'        For X = 1 To 100
'            If MapData(X, Y).autoSelect = 1 Then
'                GoTo Encontre
'            End If
'        Next
'    Next
'
'    Exit Sub
'
'Encontre:
'    Dim bX As Byte, bY As Byte, tX As Integer, tY As Integer, dif As Byte
'
'    bX = X
'    bY = Y
'
'    If MapData(X + 1, Y).autoSelect = 1 Then
'        For X = bX To 100
'            If MapData(X, bY).autoSelect = 0 Then
'                Exit For
'            End If
'        Next X
'    End If
'
'    X = X - 1
'    bX = X
'
'    loopC = True
'
'    'Ahora estamos posicionado en el primer camino
'    Do While loopC
'
'        'Si hay limite horizontal
'        For tX = X To X - 5 Step -1
'            If MapData(tX, Y).autoSelect = 0 Then
'                dif = X - tX
'                Exit For
'            End If
'        Next tX
'
'        'Si el limite es cercano
'        If dif < 4 Then
'            bY = Y
'
'            'Nos fijamos si hay algo abajo
'            If MapData(tX, Y + 1).autoSelect = 1 Then
'                'Insertamos una seguidora
'                Insertar_Superficie CAM_HORI_S, 1, X - 4, Y
'                X = tX
'
'                'Buscamos limite vertical
'                For tY = Y To Y + 5
'                    If MapData(X, tY).autoSelect = 0 Then
'                        dif = tY - Y
'                        Exit For
'                    End If
'                Next tY
'
'                'Si el limite esta cerca
'                If dif < 4 Then
'                    If MapData(X + 1, tY).autoSelect = 1 Then
'                    End If
'
'                Else 'Sino, finalizamos
'                    Insertar_Superficie CAM_VERT_S, 1, X, Y
'
'                    GoTo proximo
'                End If
'            Else 'Sino, finalizamos
'                Insertar_Superficie CAM_HORI_LF, 1, X - 4, Y
'                GoTo Salir
'            End If
'
'            If dif = 0 Then
'                If MapData(X + 1, Y).autoSelect = 1 Then
'                    Insertar_Superficie CAM_DN_RIGHT, 1, X, Y
'                ElseIf MapData(X - 1, Y).autoSelect = 1 Then
'                    Insertar_Superficie CAM_DN_LEFT, 1, X, Y
'                Else
'                    Insertar_Superficie CAM_VERT_S, 1, X, Y
'                End If
'            End If
'        Else
'            If MapData(X - 1, Y).autoSelect = 1 Then
'                Insertar_Superficie CAM_HORI_S, 1, X - 4, Y
'            End If
'
'            X = X - 4
'        End If
'
'proximo:
'
'    Loop
'
'Salir:
'    Exit Sub
'
'End Sub
Function Insertar_Superficie(ByVal i As Integer, ByVal layer As Byte, ByVal destX As Integer, ByVal destY As Integer) As Boolean
    Dim tX As Integer, tY As Integer, despTile As Integer
        
    For tY = destY To destY + SupData(i).Height
        For tX = destX To destX + SupData(i).Width
            MapData(tX, tY).Graphic(layer).GrhIndex = CInt(Val(SupData(i).Grh) + despTile)
             
            despTile = despTile + 1
        Next
    Next
End Function


