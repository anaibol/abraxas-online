Attribute VB_Name = "modSistemaComercio"
Option Explicit

Enum eModoComercio
    Compra = 1
    Venta = 2
End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 2

Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'Makes a trade. (Buy or Sell)
    
    Dim Precio As Long
    Dim Objeto As Obj
    
    If Cantidad < 1 Or Slot < 1 Then
        Exit Sub
    End If
    
    If Modo = eModoComercio.Compra Then
    
        If Slot > MaxInvSlots Then
            Exit Sub
        End If
    
        If Cantidad > MaxInvObjs Then
            Exit Sub
        End If
        
        If NpcList(NpcIndex).Inv.Obj(Slot).Amount < 1 Then
            Exit Sub
        End If
        
        If Cantidad > NpcList(NpcIndex).Inv.Obj(Slot).Amount Then
            Exit Sub
        End If
        
        Objeto.Amount = Cantidad
        Objeto.Index = NpcList(NpcIndex).Inv.Obj(Slot).Index
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        
        Precio = ObjData(NpcList(NpcIndex).Inv.Obj(Slot).Index).Valor * Cantidad

        If UserList(UserIndex).Stats.Gld < Precio Then
            Exit Sub
        End If
        
        If Not MeterEnInventario(UserIndex, Objeto) Then
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - Precio
        Call WriteUpdateGold(UserIndex)

        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNpc, Slot, Cantidad)
        Call WriteNpcInventorySlot(UserIndex, Slot, UserList(UserIndex).flags.TargetNpc)

        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.Index).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " compró del Npc " & Objeto.Amount & " " & ObjData(Objeto.Index).Name)
        ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.Index).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " compró del Npc " & Objeto.Amount & " " & ObjData(Objeto.Index).Name)
            End If
        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.Index).Type = otLlave Then
            Call WriteVar(DatPath & "Npcs.dat", "Npc" & NpcList(NpcIndex).Numero, "obj" & Slot, Objeto.Index & "-0")
            Call logVentaCasa(UserList(UserIndex).Name & " compró " & ObjData(Objeto.Index).Name)
        End If
        
    ElseIf Modo = eModoComercio.Venta Then
        
        If Slot > 200 Then
            Slot = Slot - 200
            
            If Cantidad > UserList(UserIndex).Belt.Obj(Slot).Amount Then
                Cantidad = UserList(UserIndex).Belt.Obj(Slot).Amount
            End If
            
            Objeto.Amount = Cantidad
            Objeto.Index = UserList(UserIndex).Belt.Obj(Slot).Index
            
            If Objeto.Index = 0 Then
                Exit Sub
            
            ElseIf Objeto.Index = iORO Then
                Exit Sub
            
            ElseIf ObjData(Objeto.Index).Guild Then
                'If NpcList(NpcIndex).Name <> "SR" Then
                'Call WriteConsoleMsg(UserIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                'Call EnviarNpcBelt(UserIndex, UserList(UserIndex).flags.TargetNpc)
                'EXIT SUB
                'End If
            
            ElseIf UserList(UserIndex).Belt.Obj(Slot).Amount < 0 Or Cantidad = 0 Then
                Exit Sub
            
            ElseIf Slot < LBound(UserList(UserIndex).Belt.Obj()) Or Slot > UBound(UserList(UserIndex).Belt.Obj()) Then
                Exit Sub
            
            ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No podés vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            Call QuitarBeltItem(UserIndex, Slot, Cantidad)

        Else
            If Cantidad > UserList(UserIndex).Inv.Obj(Slot).Amount Then
                Cantidad = UserList(UserIndex).Inv.Obj(Slot).Amount
            End If
            
            Objeto.Amount = Cantidad
            Objeto.Index = UserList(UserIndex).Inv.Obj(Slot).Index
            
            If Objeto.Index = 0 Then
                Exit Sub
            
            ElseIf Objeto.Index = iORO Then
                Exit Sub
            
            ElseIf ObjData(Objeto.Index).Guild Then
                'If NpcList(NpcIndex).Name <> "SR" Then
                'Call WriteConsoleMsg(UserIndex, "Las armaduras del ejército real sólo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                'Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNpc)
                'EXIT SUB
                'End If
            
            ElseIf UserList(UserIndex).Inv.Obj(Slot).Amount < 0 Or Cantidad = 0 Then
                Exit Sub
            
            ElseIf Slot < LBound(UserList(UserIndex).Inv.Obj()) Or Slot > UBound(UserList(UserIndex).Inv.Obj()) Then
                Exit Sub
            
            ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No podés vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            Call QuitarInvItem(UserIndex, Slot, Cantidad)
        
        End If
        
        Precio = ObjData(Objeto.Index).Valor * Cantidad * 0.5
        
        UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + Precio
        
        If UserList(UserIndex).Stats.Gld > MaxOro Then
            UserList(UserIndex).Stats.Gld = MaxOro
        Else
            Call WriteUpdateGold(UserIndex)
        End If
        
        If NpcList(NpcIndex).TipoItems = ObjData(Objeto.Index).Type Or NpcList(NpcIndex).TipoItems = otCualquiera Then
            Dim NpcInventorySlot As Byte
            NpcInventorySlot = SlotNpcInv(NpcIndex, Objeto.Index, Objeto.Amount)
            
            If NpcInventorySlot <= MaxInvSlots Then 'Slot válido
                'Mete el obj en el slot
                NpcList(NpcIndex).Inv.Obj(NpcInventorySlot).Index = Objeto.Index
                NpcList(NpcIndex).Inv.Obj(NpcInventorySlot).Amount = NpcList(NpcIndex).Inv.Obj(NpcInventorySlot).Amount + Objeto.Amount
                If NpcList(NpcIndex).Inv.Obj(NpcInventorySlot).Amount > MaxInvObjs Then
                    NpcList(NpcIndex).Inv.Obj(NpcInventorySlot).Amount = MaxInvObjs
                End If
                
                Call WriteNpcInventorySlot(UserIndex, NpcInventorySlot, NpcIndex)
            End If
        End If
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07\09\07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.Index).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).Name & " vendió al Npc " & Objeto.Amount & " " & ObjData(Objeto.Index).Name)
        ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.Index).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " vendió al Npc " & Objeto.Amount & " " & ObjData(Objeto.Index).Name)
            End If
        End If
        
    End If
            
End Sub

Private Function SlotNpcInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer

    SlotNpcInv = 1
    Do Until NpcList(NpcIndex).Inv.Obj(SlotNpcInv).Index = Objeto _
      And NpcList(NpcIndex).Inv.Obj(SlotNpcInv).Amount + Cantidad <= MaxInvObjs
        
        SlotNpcInv = SlotNpcInv + 1
        If SlotNpcInv > MaxInvSlots Then
            Exit Do
        End If
    Loop
    
    If SlotNpcInv > MaxInvSlots Then
    
        SlotNpcInv = 1
        
        Do Until NpcList(NpcIndex).Inv.Obj(SlotNpcInv).Index = 0
        
            SlotNpcInv = SlotNpcInv + 1
            If SlotNpcInv > MaxInvSlots Then
                Exit Do
            End If
        Loop
        
        If SlotNpcInv <= MaxInvSlots Then
            NpcList(NpcIndex).Inv.NroItems = NpcList(NpcIndex).Inv.NroItems + 1
        End If
    End If
    
End Function
