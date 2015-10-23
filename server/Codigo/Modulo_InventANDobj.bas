Attribute VB_Name = "InvNpc"
Option Explicit
'Modulo Inv & Obj: para controlar los objetos y los inventarios.

Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True, Optional UserIndexDrop As Integer = 0) As WorldPos

On Error GoTo errhandler

    If Obj.index < 1 Then
        Exit Function
    End If
    
    If MapData(Pos.X, Pos.Y).ObjInfo.index > 0 Then
        If MapData(Pos.X, Pos.Y).ObjInfo.index = Obj.index Then
            Call MakeObj(Obj, Pos.Map, Pos.X, Pos.Y)
            TirarItemAlPiso = Pos
        Else
            Dim NuevaPos As WorldPos
            
            Call Tilelibre(Pos, NuevaPos, Obj.index, NotPirata, True)
    
            If NuevaPos.X > 0 And NuevaPos.Y > 0 Then
                Call MakeObj(Obj, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                TirarItemAlPiso = NuevaPos
            End If
        End If
    Else
        Call MakeObj(Obj, Pos.Map, Pos.X, Pos.Y)
        TirarItemAlPiso = Pos
    End If
    
    If UserIndexDrop > 0 Then
    
        Dim Sonido As Byte
    
        Select Case ObjData(Obj.index).Type
            Case otAnillo
                Sonido = SND_DROP_JEWEL
            
            Case otArmadura
                Sonido = SND_DROP_ARMOR
                
            Case otPergamino
                Sonido = SND_DROP_PERGAMINO
                
            Case otEscudo
                Sonido = SND_DROP_SHIELD
        
            Case otArma
                If ObjData(Obj.index).Proyectil Then
                    Sonido = SND_DROP_BOW
                Else
                    Sonido = SND_DROP_WEAPON
                End If
                
            Case otFlecha
                Sonido = SND_DROP_ARROW
        
            Case otGuita
                Sonido = SND_DROP_COINS
        End Select
        
        Call SendData(SendTarget.ToPCArea, UserIndexDrop, Msg_SoundFX(Sonido, Pos.X, Pos.Y))
    
    End If
    
    Exit Function
errhandler:

End Function

Public Sub NpcTirarItems(ByRef Npc As Npc, ByVal UserIndexMatador As Integer)
             
'On Error Resume Next

    With Npc
    
        Dim i As Byte
        Dim MiObj As Obj
        Dim ObjIndex As Integer
        Dim TiroOro As Boolean
        
        Dim Calculo As Long
        Dim Tira As Boolean
        
        For i = 1 To MaxNpcDrops
            
            Select Case .Lvl
                     
                Case 2

                    If i > 3 Then
                        GoTo Continue
                    End If
                    
                    Calculo = 0.1 * RandomNumber(1, 25) / i
                
                Case 3
                
                    If i = 1 Then
                        GoTo Continue
                    ElseIf i = 5 Then
                        Exit For
                    End If
                    
                    Calculo = 0.2 * RandomNumber(1, 25) / i
                    
                Case 4
                
                    If i < 3 Then
                        GoTo Continue
                    End If
                    
                    Calculo = 0.3 * RandomNumber(1, 25) / i

            End Select
            
            If .Drop(i).index = iORO Then
                Tira = (Calculo >= 1)
            Else
                Tira = (Calculo >= 0.5)
            End If
            
            If Tira Then
                If .Drop(i).index = iORO Then
                    Call TirarOroNpc(.Drop(i).Amount * MultiplicadorGld, Npc.Pos)
                
                    If Not TiroOro Then
                        Call SendData(SendTarget.ToPCArea, UserIndexMatador, Msg_SoundFX(SND_DROP_COINS, .Pos.X, .Pos.Y))
                        TiroOro = True
                    End If
                    
                Else
                    MiObj.Amount = .Drop(i).Amount
                    MiObj.index = .Drop(i).index
                    Call TirarItemAlPiso(.Pos, MiObj, False, UserIndexMatador)
                End If
            End If
        
Continue:
        Next

    End With
End Sub

Public Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean

On Error Resume Next

    Dim i As Integer
    If NpcList(NpcIndex).Inv.NroItems > 0 Then
        For i = 1 To MaxInvSlots
            If NpcList(NpcIndex).Inv.Obj(i).index = ObjIndex Then
                QuedanItems = True
                Exit Function
            End If
        Next
    End If
    QuedanItems = False
End Function

Public Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
'Devuelve la Cantidad original del obj de un npc

On Error Resume Next

    Dim ln As String, NpcFile As String
    Dim i As Integer
    
    NpcFile = DatPath & "Npcs.dat"
     
    For i = 1 To MaxInvSlots
        ln = GetVar(NpcFile, "Npc" & NpcList(NpcIndex).Numero, "Obj" & i)
        If ObjIndex = Val(ReadField(1, ln, 45)) Then
            EncontrarCant = Val(ReadField(2, ln, 45))
            Exit Function
        End If
    Next
                       
    EncontrarCant = 0

End Function

Public Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error Resume Next

    Dim i As Integer
    
    With NpcList(NpcIndex)
        .Inv.NroItems = 0
    
        For i = 1 To MaxInvSlots
               .Inv.Obj(i).index = 0
               .Inv.Obj(i).Amount = 0
        Next i
    
        .InvReSpawn = 0
    End With

End Sub

Public Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'Removes a certain amount of items from a slot of an npc's inventory

    Dim ObjIndex As Integer
    Dim iCant As Integer
    
    With NpcList(NpcIndex)
        ObjIndex = .Inv.Obj(Slot).index
    
        'Quita un Obj
        If ObjData(.Inv.Obj(Slot).index).Crucial = 0 Then
            .Inv.Obj(Slot).Amount = .Inv.Obj(Slot).Amount - Cantidad
            
            If .Inv.Obj(Slot).Amount < 1 Then
                .Inv.NroItems = .Inv.NroItems - 1
                .Inv.Obj(Slot).index = 0
                .Inv.Obj(Slot).Amount = 0
                If .Inv.NroItems = 0 And .InvReSpawn <> 1 Then
                    Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
        Else
            .Inv.Obj(Slot).Amount = .Inv.Obj(Slot).Amount - Cantidad
            
            If .Inv.Obj(Slot).Amount < 1 Then
                .Inv.NroItems = .Inv.NroItems - 1
                .Inv.Obj(Slot).index = 0
                .Inv.Obj(Slot).Amount = 0
                
                If Not QuedanItems(NpcIndex, ObjIndex) Then
                    'Check if the item is in the npc's dat.
                    iCant = EncontrarCant(NpcIndex, ObjIndex)
                    If iCant Then
                        .Inv.Obj(Slot).index = ObjIndex
                        .Inv.Obj(Slot).Amount = iCant
                        .Inv.NroItems = .Inv.NroItems + 1
                    End If
                End If
                
                If .Inv.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
        End If
    End With
End Sub

Public Sub CargarInvent(ByVal NpcIndex As Integer)

    'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC As Integer
    Dim ln As String
    Dim NpcFile As String
    
    NpcFile = DatPath & "Npcs.dat"
    
    With NpcList(NpcIndex)
        .Inv.NroItems = Val(GetVar(NpcFile, "Npc" & .Numero, "NROItemS"))
        
        For LoopC = 1 To .Inv.NroItems
            ln = GetVar(NpcFile, "Npc" & .Numero, "Obj" & LoopC)
            .Inv.Obj(LoopC).index = Val(ReadField(1, ln, 45))
            .Inv.Obj(LoopC).Amount = Val(ReadField(2, ln, 45))
        Next LoopC
    End With

End Sub

Public Sub TirarOroNpc(ByVal Cantidad As Long, ByRef Pos As WorldPos)

On Error GoTo errhandler

    If Cantidad > 0 Then
        Dim MiObj As Obj
        
        MiObj.Amount = Cantidad
        MiObj.index = iORO
            
        Call TirarItemAlPiso(Pos, MiObj)
    End If

    Exit Sub

errhandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & ": " & Err.description)
End Sub


