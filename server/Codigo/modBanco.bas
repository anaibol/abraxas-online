Attribute VB_Name = "modBanco"
Option Explicit

Public Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Byte, ByVal Cantidad As Integer)

    Dim Slot As Byte
    Dim obji As Integer
    
    With UserList(UserIndex)
        If .Bank.Obj(ObjIndex).Amount < 1 Then
            Exit Sub
        End If
        
        obji = .Bank.Obj(ObjIndex).Index
        
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until .Inv.Obj(Slot).Index = obji And _
           .Inv.Obj(Slot).Amount + Cantidad <= MaxInvObjs
            
            Slot = Slot + 1
            If Slot > MaxInvSlots Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio
        If Slot > MaxInvSlots Then
            Slot = 1
            Do Until .Inv.Obj(Slot).Index = 0
                Slot = Slot + 1
    
                If Slot > MaxInvSlots Then
                    Call WriteConsoleMsg(UserIndex, "No podés llevar nada más.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop
            .Inv.NroItems = .Inv.NroItems + 1
        End If
        
        'Mete el obj en el slot
        If .Inv.Obj(Slot).Amount + Cantidad <= MaxInvObjs Then
            'Menor que MaxInvObjs
            .Inv.Obj(Slot).Index = obji
            .Inv.Obj(Slot).Amount = .Inv.Obj(Slot).Amount + Cantidad
            
            Call QuitarBancoInvItem(UserIndex, ObjIndex, Cantidad)
        
            'Actualizamos el inventario del usuario
            Call WriteInventorySlot(UserIndex, Slot)
                        
            'Actualizamos el banco
            Call WriteBankSlot(UserIndex, ObjIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No podés llevar nada más.", FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
        
End Sub

Public Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    
    Dim ObjIndex As Integer
    
    With UserList(UserIndex)
        ObjIndex = .Bank.Obj(Slot).Index
    
        .Bank.Obj(Slot).Amount = .Bank.Obj(Slot).Amount - Cantidad
        
        If .Bank.Obj(Slot).Amount < 1 Then
            .Bank.Obj(Slot).Index = 0
            .Bank.Obj(Slot).Amount = 0
            
            .Bank.NroItems = .Bank.NroItems - 1
            
            Call WriteBankSlot(UserIndex, Slot)
        End If
    End With
    
End Sub

Public Sub UserDejaObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

    Dim BankSlot As Byte
    Dim ObjIndex As Integer
    
    If Cantidad < 1 Then
        Exit Sub
    End If
    
    With UserList(UserIndex)
        If Slot > 200 Then
            ObjIndex = .Belt.Obj(Slot - 200).Index
        Else
            ObjIndex = .Inv.Obj(Slot).Index
        End If
    
        '¿Ya tiene un objeto de este tipo?
        BankSlot = 1
        
        Do Until .Bank.Obj(BankSlot).Index = ObjIndex And _
            .Bank.Obj(BankSlot).Amount + Cantidad <= MaxInvObjs
            BankSlot = BankSlot + 1
            
            If BankSlot > MaxBankSlots Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If BankSlot > MaxBankSlots Then
            BankSlot = 1
            Do Until .Bank.Obj(BankSlot).Index = 0
                BankSlot = BankSlot + 1
                
                If BankSlot > MaxBankSlots Then
                    Call WriteConsoleMsg(UserIndex, "No tenés más espacio en el banco.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop
            
            .Bank.NroItems = .Bank.NroItems + 1
        End If
        
        If BankSlot <= MaxBankSlots Then 'BankSlot válido
            'Mete el obj en el slot
            If .Bank.Obj(BankSlot).Amount + Cantidad <= MaxInvObjs Then
                
                'Menor que MaxInvObjs
                .Bank.Obj(BankSlot).Index = ObjIndex
                .Bank.Obj(BankSlot).Amount = .Bank.Obj(BankSlot).Amount + Cantidad
                        
                If Slot > 200 Then
                    Call QuitarBeltItem(UserIndex, Slot - 200, Cantidad)
                Else
                    Call QuitarInvItem(UserIndex, Slot, Cantidad)
                End If
                
                'Actualizamos el inventario del banco
                Call WriteBankSlot(UserIndex, BankSlot)
            Else
                Call WriteConsoleMsg(UserIndex, "No tenés más lugar en la boveda.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

Public Sub SendUserBovedaTxt(ByVal SendIndex As Integer, ByVal UserIndex As Integer)

On Error Resume Next

    Dim j As Byte
    
    Call WriteConsoleMsg(SendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(SendIndex, "Tiene " & UserList(UserIndex).Bank.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To MaxBankSlots
        If UserList(UserIndex).Bank.Obj(j).Index > 0 Then
            Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).Bank.Obj(j).Index).Name & " Cantidad:" & UserList(UserIndex).Bank.Obj(j).Amount, FontTypeNames.FONTTYPE_INFO)
        End If
    Next

End Sub

Public Sub SendUserBovedaTxtFromChar(ByVal SendIndex As Integer, ByVal CharName As String)

On Error Resume Next

    Dim BankStr As String
    Dim TempStr() As String
    Dim TempStr2() As String
    
    Dim i As Byte
    
    Dim Tmp As String
    Dim ObjInd As Long, ObjCant As Long
            
    If User_Exist(CharName) Then
    
        Call DB_RS_Open("SELECT * FROM people WHERE `name`='" & CharName & "'")
        
        BankStr = DB_RS!Bank
        
        DB_RS.Close
        
        'Inventory string
        If LenB(BankStr) > 0 Then
            TempStr = Split(BankStr, vbNewLine)  'Split up the inventory slots
            For i = 0 To MaxBankSlots        'Loop through the slots
                TempStr2 = Split(TempStr(i), " ")   'Split up the slot, objindex, amount and equipted (in that order)
                If Val(TempStr2(0)) <= MaxInvSlots Then
                    'With .inv.Obj(val(TempStr2(0)))
                    '.index = val(TempStr2(1))
                    '.Amount = val(TempStr2(2))
                    '.Equipped = val(TempStr2(3))
                    'End With
                    '.inv.NroItems = .inv.NroItems + 1
                End If
            Next i
        End If
        
        Dim Nro As Byte
        
        Call WriteConsoleMsg(SendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        
        'Nro = GetVar(CharFile, "BancoInventory", "CantidadItems")
        
        Call WriteConsoleMsg(SendIndex, "Tiene " & Nro & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For i = 1 To Nro
            'Tmp = GetVar(CharFile, "BancoInventory", "Obj" & i)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(SendIndex, "Objeto " & i & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next
    Else
        Call WriteConsoleMsg(SendIndex, "Personaje inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

