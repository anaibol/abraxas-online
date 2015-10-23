Attribute VB_Name = "InvUsuario"
Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
    
On Error Resume Next
    
    Dim i As Byte
    Dim ObjIndex As Integer
    
    For i = 1 To MaxInvSlots
        ObjIndex = UserList(UserIndex).Inv.Obj(i).index
        If ObjIndex > 0 Then
            If ItemSeCae(ObjIndex) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
        End If
    Next i

End Function

Public Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

On Error GoTo manejador
    
    Dim flag As Boolean
    
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) > 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                    Exit Function
                End If
            Next i
        End If
    End If
    
    ClasePuedeUsarItem = True
    
    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Public Sub ResetUserInventario(ByVal UserIndex As Integer)

    Dim j As Integer
            
    With UserList(UserIndex)
        For j = 1 To MaxInvSlots
            .Inv.Obj(j).index = 0
            .Inv.Obj(j).Amount = 0
        Next j
    
        .Inv.NroItems = 0
    
        .Inv.Body = 0
        .Inv.Head = 0
        .Inv.LeftHand = 0
        .Inv.RightHand = 0
        .Inv.AmmoAmount = 0
        .Inv.Belt = 0
        .Inv.Ring = 0
        .Inv.Ship = 0
    End With

End Sub

Public Sub ResetUserCinturon(ByVal UserIndex As Integer)
    Dim j As Integer
            
    With UserList(UserIndex)
        For j = 1 To MaxBeltSlots
            .Belt.Obj(j).index = 0
            .Belt.Obj(j).Amount = 0
        Next j
    
        .Belt.NroItems = 0
    End With
End Sub

Public Sub ResetUserHechizos(ByVal UserIndex As Integer)
    UserList(UserIndex).Spells.Nro = 0
    
    Dim LoopC As Byte
    For LoopC = 1 To MaxSpellSlots
        UserList(UserIndex).Spells.Spell(LoopC) = 0
    Next LoopC
End Sub

Public Sub ResetUserCompanieros(ByVal UserIndex As Integer)
    UserList(UserIndex).Compas.Nro = 0

    Dim LoopC As Byte
    For LoopC = 1 To MaxCompaSlots
        UserList(UserIndex).Compas.Compa(LoopC) = vbNullString
    Next LoopC
    
End Sub

Public Sub ResetUserMascotas(ByVal UserIndex As Integer)

    UserList(UserIndex).Pets.Nro = 0
    UserList(UserIndex).Pets.NroALaVez = 0

    Dim LoopC As Byte
        
    For LoopC = 1 To MaxPets
    
        With UserList(UserIndex).Pets.Pet(LoopC)
            .index = 0
            .Tipo = 0
            .Lvl = 0
            .Elu = 0
            .Exp = 0
            .MinHP = 0
            .MaxHP = 0
            .MaxHP = 0
            .MaxHP = 0
            .MinHit = 0
            .MaxHit = 0
            .Def = 0
            .DefM = 0
        End With
        
    Next LoopC
    
End Sub

Public Sub ResetUserBanco(ByVal UserIndex As Integer)
    
    UserList(UserIndex).Bank.NroItems = 0
        
    Dim LoopC As Long
    
    For LoopC = 1 To MaxBankSlots
            
        With UserList(UserIndex).Bank.Obj(LoopC)
          .index = 0
          .Amount = 0
        End With
        
    Next LoopC
    
    UserList(UserIndex).Bank.NroItems = 0
End Sub

Public Sub ResetUserPlataformas(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MaxPlataformSlots
            
        With UserList(UserIndex).Plataformas.Plataforma(LoopC)
          .Map = 0
          .X = 0
          .Y = 0
        End With
        
    Next LoopC
    
    UserList(UserIndex).Plataformas.Nro = 0
End Sub

Public Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

On Error GoTo errhandler

    With UserList(UserIndex)
    
        If Cantidad > 0 And Cantidad <= .Stats.Gld Then
        
            Dim i As Byte
            Dim MiObj As Obj
            'info debug
            Dim Loops As Integer
            
            If Cantidad > 100000 Then
                
                Dim j As Integer
                Dim k As Integer
                Dim M As Integer
                Dim Cercanos As String
                
                M = .Pos.Map
                
                For j = .Pos.X - 10 To .Pos.X + 10
                    For k = .Pos.Y - 10 To .Pos.Y + 10
                        If InMapBounds(M, j, k) Then
                            If MapData(j, k).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(j, k).UserIndex).Name & ","
                            End If
                        End If
                    Next k
                Next j
                
                Call LogDesarrollo(.Name & " tiró " & Cantidad & " monedas de oro. Cercanos: " & Cercanos)
                
            End If
    
            If EsGM(UserIndex) Then
                Call LogGM(.Name, "Tiró " & Cantidad & " monedas de oro. Cercanos: " & Cercanos)
            End If
            
            MiObj.index = iORO
            MiObj.Amount = Cantidad
            
            Dim AuxPos As WorldPos
            
            If .Clase = eClass.Pirat And .Inv.Ship = 476 Then
                AuxPos = TirarItemAlPiso(.Pos, MiObj, False, UserIndex)
                
                If AuxPos.X > 0 And AuxPos.Y > 0 Then
                    .Stats.Gld = .Stats.Gld - Cantidad
                End If
                
            Else
                AuxPos = TirarItemAlPiso(.Pos, MiObj, , UserIndex)
                
                If AuxPos.X > 0 And AuxPos.Y > 0 Then
                    .Stats.Gld = .Stats.Gld - Cantidad
                End If
            End If
            
        End If
        
    End With
    
    Exit Sub

errhandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & ": " & Err.description)
End Sub

Public Sub QuitarInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal Cantidad As Integer = 1)
    
    If Slot < 1 Or Slot > MaxInvSlots Then
        Exit Sub
    End If
    
    With UserList(UserIndex).Inv.Obj(Slot)
        .Amount = .Amount - Cantidad
        
        If .Amount < 1 Then
            .Amount = 0
            .index = 0

            UserList(UserIndex).Inv.NroItems = UserList(UserIndex).Inv.NroItems - 1
            
            Slot = Slot + 200
            
            Call WriteSlotMenosUno(UserIndex, Slot)
        
        ElseIf Cantidad = 1 Then
            Call WriteSlotMenosUno(UserIndex, Slot)
        Else
            Call WriteInventorySlot(UserIndex, Slot)
        End If
    End With

End Sub

Public Sub QuitarBeltItem(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal Cantidad As Integer = 1)

On Error GoTo errhandler
    If Slot < 1 Or Slot > MaxBeltSlots Then
        Exit Sub
    End If
    
    With UserList(UserIndex).Belt.Obj(Slot)
        .Amount = .Amount - Cantidad
        
        If .Amount < 1 Then
            .Amount = 0
            .index = 0

            UserList(UserIndex).Belt.NroItems = UserList(UserIndex).Belt.NroItems - 1
        End If
        
        Call WriteBeltSlot(UserIndex, Slot)
    End With
Exit Sub

errhandler:
    Call LogError("Error en QuitarInvItem. Error " & Err.Number & ": " & Err.description)
    
End Sub

Public Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer)

    Dim Obj As Obj
    
    With UserList(UserIndex)
    
        If Num > 0 And Num <= MaxInvObjs Then
        
            If Num > .Inv.Obj(Slot).Amount Then
                Num = .Inv.Obj(Slot).Amount
            End If
          
            Obj.index = .Inv.Obj(Slot).index
          
            Obj.Amount = Num
            
            Call TirarItemAlPiso(.Pos, Obj, , UserIndex)
                
            Call QuitarInvItem(UserIndex, Slot, Num)
                
            If ObjData(Obj.index).Type = otBarco Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
            End If
            
            If Not .flags.Privilegios And PlayerType.User Then
                Call LogGM(.Name, "Tiro Cantidad:" & Num & " Objeto:" & ObjData(Obj.index).Name)
            End If
            
            'Log de Objetos que se tiran al piso
            'Es un Objeto que tenemos que loguear?
            If ObjData(Obj.index).Log = 1 Then
                Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.index).Name & " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y)
            ElseIf Obj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(Obj.index).NoLog <> 1 Then
                    Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.index).Name & " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y)
                End If
            End If
        End If
    End With
End Sub

Public Sub DropBeltObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer)

    Dim Obj As Obj
    
    With UserList(UserIndex)
    
        If Num > 0 And Num <= MaxBeltObjs Then
        
            If Num > .Belt.Obj(Slot).Amount Then
                Num = .Belt.Obj(Slot).Amount
            End If
          
            Obj.index = .Belt.Obj(Slot).index
          
            Obj.Amount = Num
            
            Call TirarItemAlPiso(.Pos, Obj, , UserIndex)
                
            Call QuitarBeltItem(UserIndex, Slot, Num)
            
            'Log de Objetos que se tiran al piso
            'Es un Objeto que tenemos que loguear?
            If ObjData(Obj.index).Log = 1 Then
                Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.index).Name & " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y)
            ElseIf Obj.Amount > 5000 Then 'Es mucha cantidad? > Subí a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(Obj.index).NoLog <> 1 Then
                    Call LogDesarrollo(.Name & " tiró al piso " & Obj.Amount & " " & ObjData(Obj.index).Name & " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y)
                End If
            End If
        End If
    End With
End Sub

Public Sub EraseObj(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Num As Long = -1)
    With MapData(X, Y)
    
        If .ObjInfo.Amount > 0 Then
            If Num = -1 Then
                .ObjInfo.Amount = 0
            Else
                .ObjInfo.Amount = .ObjInfo.Amount - Num
            End If
            
            If .ObjInfo.Amount < 1 Then
                .ObjInfo.Amount = 0
                .ObjInfo.index = 0
                
                Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjDelete(X, Y))
            End If
            
            'If .ObjInfo.index = iObjCuerpoMuerto Then
            '    If .Blocked Then
            '    .Blocked = False
            '    End If
            'End If
            
        Else
            .ObjInfo.index = 0
            .ObjInfo.Amount = 0
            
            Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjDelete(X, Y))
        End If
    
    End With
End Sub

Public Sub MakeObj(ByRef Obj As Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    
    If Obj.index > 0 Then
        If Obj.index <= UBound(ObjData) Then
        
            With MapData(X, Y)
            
                If .ObjInfo.index > 0 Then
                    If .ObjInfo.index = Obj.index Then
                        If .ObjInfo.Amount > 0 Then
                            .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
                        End If
                    Else
                        .ObjInfo = Obj
                    End If
                Else
                    .ObjInfo = Obj
                End If
                    
                If ObjData(.ObjInfo.index).Type = otPortal Then
                    Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, X, Y, , 10000 + .TileExit.Map))
                    
                ElseIf ObjData(.ObjInfo.index).Type = otAlijo Then
                    Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, X, Y, ObjData(.ObjInfo.index).Name, 1))
                                
                ElseIf ObjData(.ObjInfo.index).Type = otCuerpoMuerto Then
                    If .ObjInfo.Amount > 0 Then
                        If UserList(.ObjInfo.Amount).Stats.Muerto Then
                            Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, X, Y, UserList(.ObjInfo.Amount).Name))
                        Else
                            Call EraseObj(Map, X, Y, .ObjInfo.Amount)
                        End If
                    Else
                        Call EraseObj(Map, X, Y, -1)
                    End If
             
                ElseIf ObjData(.ObjInfo.index).Type = otGuita Then
                    Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjCreate(0, ObjData(.ObjInfo.index).Type, X, Y, , .ObjInfo.Amount))

                Else
                    Call modSendData.SendToAreaByPos(Map, X, Y, Msg_ObjCreate(ObjData(.ObjInfo.index).GrhIndex, ObjData(.ObjInfo.index).Type, X, Y, ObjData(.ObjInfo.index).Name, .ObjInfo.Amount))
                End If
                
            End With
        End If
    End If
End Sub

Public Function MeterEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj, Optional ByVal Update = True) As Boolean
On Error GoTo errhandler

    'Call LogTarea("MeterEnInventario")

    Dim X As Integer
    Dim Y As Integer
    Dim Slot As Byte
    
    With UserList(UserIndex)
    
        If .Inv.NroItems > 0 Then
            For Slot = 1 To MaxInvSlots
                If .Inv.Obj(Slot).index = MiObj.index And .Inv.Obj(Slot).Amount < MaxInvObjs Then
                    Dim Amount As Long
                    
                    Amount = .Inv.Obj(Slot).Amount + MiObj.Amount

                    If Amount > MaxInvObjs Then
                        MiObj.Amount = Amount - MaxInvObjs
                        .Inv.Obj(Slot).Amount = MaxInvObjs
                        
                        If Update Then
                            Call WriteInventorySlot(UserIndex, Slot)
                        End If
                    
                    Else
                        .Inv.Obj(Slot).Amount = Amount
                        MiObj.Amount = 0
                        
                        If Update Then
                            Call WriteInventorySlot(UserIndex, Slot)
                        End If
                        
                        Exit For
                    End If
                End If
            Next Slot
        End If
        
        If MiObj.Amount > 0 Then
            For Slot = 1 To MaxInvSlots
                If .Inv.Obj(Slot).index = 0 Then
                    .Inv.Obj(Slot).index = MiObj.index
                    .Inv.Obj(Slot).Amount = MiObj.Amount

                    If .Inv.Obj(Slot).Amount > MaxInvObjs Then
                        MiObj.Amount = .Inv.Obj(Slot).Amount - MaxInvObjs
                        .Inv.Obj(Slot).Amount = MaxInvObjs
                        
                        If Update Then
                            Call WriteInventorySlot(UserIndex, Slot)
                        End If
                    
                    Else
                        MiObj.Amount = 0
                        
                        If Update Then
                            Call WriteInventorySlot(UserIndex, Slot)
                        End If
                        
                        Exit For
                    End If
                End If
            Next Slot
        End If

        If MiObj.Amount > 0 Then
            Call WriteConsoleMsg(UserIndex, "No podés llevar nada más.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        .Inv.NroItems = .Inv.NroItems + 1
    
        MeterEnInventario = True
    
        'If .flags.Desnudo And ObjData(MiObj.index).Type = otArmadura Then
        '    Call Equipar(UserIndex, Slot)
        'End If

    End With
    
    Exit Function
errhandler:
    Call LogError("Error en MeterEnInventario. Error " & Err.Number & ": " & Err.description)
End Function

Public Function MeterEnCinturon(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

    'Call LogTarea("MeterEnCinturon")

    Dim X As Integer
    Dim Y As Integer
    Dim Slot As Byte
    
    If ObjData(MiObj.index).Type <> otPocion Then
        Exit Function
    End If
    
    With UserList(UserIndex)
    
        If .Belt.NroItems > 0 Then
            For Slot = 1 To MaxBeltSlots
                If .Belt.Obj(Slot).index = MiObj.index And .Belt.Obj(Slot).Amount < MaxBeltObjs Then
                    .Belt.Obj(Slot).Amount = .Belt.Obj(Slot).Amount + MiObj.Amount
                    
                    If .Belt.Obj(Slot).Amount > MaxBeltObjs Then
                        MiObj.Amount = .Belt.Obj(Slot).Amount - MaxBeltObjs
                        .Belt.Obj(Slot).Amount = MaxBeltObjs
                        Call WriteBeltSlot(UserIndex, Slot)
                    Else
                        MiObj.Amount = 0
                        Call WriteBeltSlot(UserIndex, Slot)
                        Exit For
                    End If
                End If
            Next Slot
        End If
        
        If MiObj.Amount > 0 Then
            For Slot = 1 To MaxBeltSlots
                If .Belt.Obj(Slot).index = 0 Then
                    .Belt.Obj(Slot).index = MiObj.index
                    .Belt.Obj(Slot).Amount = MiObj.Amount
                
                    If .Belt.Obj(Slot).Amount > MaxBeltObjs Then
                        MiObj.Amount = .Belt.Obj(Slot).Amount - MaxBeltObjs
                        .Belt.Obj(Slot).Amount = MaxBeltObjs
                        Call WriteBeltSlot(UserIndex, Slot)
                    Else
                        MiObj.Amount = 0
                        Call WriteBeltSlot(UserIndex, Slot)
                        Exit For
                    End If
                End If
            Next Slot
        End If

        If MiObj.Amount > 0 Then
            Call MeterEnInventario(UserIndex, MiObj)
        End If
        
        .Belt.NroItems = .Belt.NroItems + 1
        
        MeterEnCinturon = True

    End With
    
    Exit Function
errhandler:
    Call LogError("Error en MeterEnCinturon. Error " & Err.Number & ": " & Err.description)
End Function

Public Sub GetObj(ByVal UserIndex As Integer)

    Dim Obj As ObjData
    Dim MiObj As Obj
    Dim ObjPos As String
    
    With UserList(UserIndex)
        '¿Hay algun obj?
        If MapData(.Pos.X, .Pos.Y).ObjInfo.index > 0 Then
        
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.X, .Pos.Y).ObjInfo.index).Agarrable Then
            
                Dim X As Integer
                Dim Y As Integer
                Dim Slot As Byte
            
                X = .Pos.X
                Y = .Pos.Y
                Obj = ObjData(MapData(X, Y).ObjInfo.index)
                MiObj.Amount = MapData(X, Y).ObjInfo.Amount
                MiObj.index = MapData(X, Y).ObjInfo.index
            
                If ObjData(MiObj.index).Type = otGuita Then
                
                    UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + MiObj.Amount
                    Call EraseObj(UserList(UserIndex).Pos.Map, X, Y, -1)
                                        
                    Call WriteUpdateGold(UserIndex)
                Else
                    If Not MeterEnCinturon(UserIndex, MiObj) Then
                        If Not MeterEnInventario(UserIndex, MiObj) Then
                            Exit Sub
                        End If
                    End If
                
                    'Quitamos el objeto
                    Call EraseObj(.Pos.Map, X, Y, MapData(X, Y).ObjInfo.Amount)
                    
                    If Not .flags.Privilegios And PlayerType.User Then
                        Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.index).Name)
                    End If
                
                    'Log de Objetos que se agarran del piso.
                    'Es un Objeto que tenemos que loguear?
                    If ObjData(MiObj.index).Log = 1 Then
                        ObjPos = " Mapa: " & .Pos.Map & " X: " & X & " Y: " & Y
                        Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.index).Name & ObjPos)
                    ElseIf MiObj.Amount >= MaxInvObjs - 1000 Then 'Es mucha cantidad?
                        'Si no es de los prohibidos de loguear, lo logueamos.
                        If ObjData(MiObj.index).NoLog <> 1 Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & X & " Y: " & Y
                            Call LogDesarrollo(.Name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.index).Name & ObjPos)
                        End If
                    End If
                End If
            End If
        End If
        
    End With
End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal ObjType As eObjType)

    Dim Obj As Obj
    
    Obj.Amount = 1
    
    With UserList(UserIndex)
    
        Select Case ObjType
        
            Case otArma
            
                If UsaArco(UserIndex) > 0 Then
                    Obj.index = .Inv.LeftHand
                    .Inv.LeftHand = 0
                Else
                    Obj.index = .Inv.RightHand
                    .Inv.RightHand = 0
                End If
                
                If Not .flags.Mimetizado Then
                    .Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                End If
                
            Case otAnillo
                Obj.index = .Inv.Ring
                .Inv.Ring = 0
            
            Case otFlecha
                Obj.index = .Inv.RightHand
                Obj.Amount = .Inv.AmmoAmount
                
                .Inv.RightHand = 0
                .Inv.AmmoAmount = 0
            
            Case otArmadura
                Obj.index = .Inv.Body
                .Inv.Body = 0
                
                Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado)
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
            
            Case otCasco
                Obj.index = .Inv.Head
                .Inv.Head = 0
            
                If Not .flags.Mimetizado Then
                    .Char.HeadAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                End If
                
            Case otEscudo
                Obj.index = .Inv.LeftHand
                .Inv.LeftHand = 0
        
                If Not .flags.Mimetizado Then
                    .Char.ShieldAnim = NingunEscudo
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                End If
                
            Case otCinturon
                Obj.index = .Inv.Belt
                .Inv.Belt = 0
                
                Call ResetUserCinturon(UserIndex)
        
        End Select
        
        Call MeterEnInventario(UserIndex, Obj, False)

    End With
        
End Sub

Public Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
End Function

Public Function GuildaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    If ObjData(ObjIndex).Guild = 1 Then
        GuildaPuedeUsarItem = False
    Else
        GuildaPuedeUsarItem = True
    End If
End Function

Public Sub Equipar(ByVal UserIndex As Integer, ByVal Slot As Byte)

    Dim Obj As ObjData
    Dim ObjIndex As Integer
    Dim Equipped  As Boolean
    
    With UserList(UserIndex)
        ObjIndex = .Inv.Obj(Slot).index
        Obj = ObjData(ObjIndex)
           
        Select Case Obj.Type
        
            Case otArma
            
               If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                  GuildaPuedeUsarItem(UserIndex, ObjIndex) Then

                    If UsaArco(UserIndex) > 0 Or UsaArmaNoArco(UserIndex) > 0 Then
                        Call Desequipar(UserIndex, otArma)
                    End If

                    If Obj.Proyectil Then
                        .Inv.LeftHand = ObjIndex
                    Else
                        .Inv.RightHand = ObjIndex
                    End If
                    
                    Equipped = True
                                        
                    'El sonido solo se envia si no lo produce un admin invisible
                    If .flags.AdminInvisible < 1 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_SACARARMA, .Pos.X, .Pos.Y))
                    End If
                    
                    If .flags.Mimetizado Then
                        .CharMimetizado.WeaponAnim = Obj.WeaponAnim
                    Else
                        .Char.WeaponAnim = Obj.WeaponAnim
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                    End If
               End If
        
        Case otAnillo
            
            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
            GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                
                If .Inv.Ring > 0 Then
                    Call Desequipar(UserIndex, otAnillo)
                End If
                
                .Inv.Ring = ObjIndex
                Equipped = True
            End If
        
        Case otFlecha
            
            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                
                If .Inv.AmmoAmount > 0 Then
                    Call Desequipar(UserIndex, otFlecha)
                End If
                
                .Inv.RightHand = ObjIndex
                .Inv.AmmoAmount = .Inv.Obj(Slot).Amount
                Equipped = True
            End If
        
        Case otArmadura

            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                SexoPuedeUsarItem(UserIndex, ObjIndex) And _
                CheckRazaUsaRopa(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
            
                If .Inv.Body > 0 Then
                    Call Desequipar(UserIndex, otArmadura)
                End If
        
                .Inv.Body = ObjIndex
                Equipped = True
                        
                If .flags.Mimetizado Then
                    .CharMimetizado.Body = Obj.BodyAnim
                Else
                    .Char.Body = Obj.BodyAnim
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                End If
                
                .flags.Desnudo = False
            End If
        
        Case otCasco
                
                If ClasePuedeUsarItem(UserIndex, ObjIndex) Then

                    If .Inv.Head > 0 Then
                        Call Desequipar(UserIndex, otCasco)
                    End If
                    
                    .Inv.Head = ObjIndex
                    Equipped = True
                    
                    If .flags.Mimetizado Then
                        .CharMimetizado.HeadAnim = Obj.HeadAnim
                    Else
                        .Char.HeadAnim = Obj.HeadAnim
                        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                    End If
                End If
        
        Case otEscudo
            
            If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
                GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
    
                If UsaEscudo(UserIndex) > 0 Then
                    Call Desequipar(UserIndex, otEscudo)
                End If
    
                .Inv.LeftHand = ObjIndex
                Equipped = True
                
                If .flags.Mimetizado Then
                    .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                Else
                    .Char.ShieldAnim = Obj.ShieldAnim
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
                End If
            End If
            
        End Select
        
        If Equipped Then
            If .Inv.Obj(Slot).Amount > 1 And Obj.Type <> otFlecha Then
                .Inv.Obj(Slot).Amount = .Inv.Obj(Slot).Amount - 1
            Else
                .Inv.Obj(Slot).index = 0
                .Inv.Obj(Slot).Amount = 0
            End If
        End If
        
    End With
        
End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

    With UserList(UserIndex)
        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or _
           .Raza = eRaza.Elfo Or _
           .Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False
        End If
    End With

    Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Public Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Handels the usage of Items from inventory box.

    Dim Obj As ObjData
    Dim ObjIndex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj
    
    With UserList(UserIndex)
    
    Obj = ObjData(.Inv.Obj(Slot).index)
    
    If Obj.Proyectil Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then
            Exit Sub
        End If
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then
            Exit Sub
        End If
    End If
    
    ObjIndex = .Inv.Obj(Slot).index
    .flags.TargetObjInvIndex = ObjIndex
    .flags.TargetObjInvSlot = Slot
    
    Select Case Obj.Type
    
        Case otArma
            
            If .Stats.MinSta < 1 Then
                Call WriteConsoleMsg(UserIndex, "No tenés energía", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.TargetObjIndex > 0 Then
                If .flags.TargetObjIndex = Leña Then
                    If .Inv.Obj(Slot).index = DAGA Then
                        Call TratarDeHacerFogata(.flags.TargetObjMap, _
                            .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                    End If
                
                ElseIf ObjData(.flags.TargetObjIndex).Type = otFragua Then 'fragua => TODO: hacer una constante para el Index de la fragua
                    If ObjData(.Inv.Obj(Slot).index).Type = otArma Then
                        Call FundirArmas(UserIndex)
                    End If
                End If
                
            ElseIf ObjIndex = SERRUCHO_CARPINTERO Then
                Call WriteCarpenterObjs(UserIndex)
            End If
    
        Case otUseOnce
        
            'Usa el Item
            .Stats.MinHam = .Stats.MinHam + Obj.MinHam
            
            If .Stats.MinHam > 100 Then
                .Stats.MinHam = 100
            End If
            
            Call WriteUpdateHungerAndThirst(UserIndex)
            
            'Sonido
            If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Then
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
            Else
                Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
            End If
            
            'Quitamos del inv el Item
            Call QuitarInvItem(UserIndex, Slot)
    
        Case otPocion
                
            If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                Exit Sub
            End If
            
            .flags.TipoPocion = Obj.TipoPocion
                    
            Select Case .flags.TipoPocion
            
                Case 1 'Modif la agilidad
                    .flags.DuracionEfecto = Obj.DuracionEfecto

                    If .Stats.Atributos(eAtributos.Agilidad) < 2 * .Stats.AtributosBackUP(Agilidad) Then
                        'Usa el Item
                        .Stats.Atributos(eAtributos.Agilidad) = .Stats.Atributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        
                        If .Stats.Atributos(eAtributos.Agilidad) > MaxAtributos Then
                            .Stats.Atributos(eAtributos.Agilidad) = MaxAtributos
                        End If
                        
                        If .Stats.Atributos(eAtributos.Agilidad) > 2 * .Stats.AtributosBackUP(Agilidad) Then
                            .Stats.Atributos(eAtributos.Agilidad) = 2 * .Stats.AtributosBackUP(Agilidad)
                        End If
                        
                        Call WriteUpdateDexterity(UserIndex)
                    End If

                    'Quitamos del inv el Item
                    Call QuitarInvItem(UserIndex, Slot)
                    
                    'Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible > 0 Then
                        Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    End If

                Case 2 'Modif la fuerza
                    .flags.DuracionEfecto = Obj.DuracionEfecto
                
                    If .Stats.Atributos(eAtributos.Fuerza) < 2 * .Stats.AtributosBackUP(Fuerza) Then
                        .Stats.Atributos(eAtributos.Fuerza) = .Stats.Atributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)

                        If .Stats.Atributos(eAtributos.Fuerza) > MaxAtributos Then
                            .Stats.Atributos(eAtributos.Fuerza) = MaxAtributos
                        End If
                        
                        If .Stats.Atributos(eAtributos.Fuerza) > 2 * .Stats.AtributosBackUP(Fuerza) Then
                            .Stats.Atributos(eAtributos.Fuerza) = 2 * .Stats.AtributosBackUP(Fuerza)
                        End If
                        
                        Call WriteUpdateStrenght(UserIndex)
                    End If
                
                    'Quitamos del inv el Item
                    Call QuitarInvItem(UserIndex, Slot)
                    
                    'Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible > 0 Then
                        Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    End If
                    
                Case 3 'Pocion roja, restaura HP
                
                    'Usa el Item SI NO TIENE LA VIDA COMPLETA
                    If .Stats.MinHP < .Stats.MaxHP Then
                        .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        Call WriteUpdateHP(UserIndex)
                        
                        'Quitamos del inv el Item
                        Call QuitarInvItem(UserIndex, Slot)
                        
                        'Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible > 0 Then
                            Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        End If
                    End If
                    
                Case 4 'Pocion azul, restaura MANA
                
                    'Usa el Item SI NO TIENE LA MANÁ COMPLETA
                    If .Stats.MinMan < .Stats.MaxMan Then
                        .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 5) + .Stats.Elv * 0.33 + 50 \ .Stats.Elv
                        Call WriteUpdateMana(UserIndex)
        
                        'Quitamos del inv el Item
                        Call QuitarInvItem(UserIndex, Slot)
                        
                        'Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible > 0 Then
                            Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        End If
                    End If
                    
                Case 5 'Pocion violeta
                
                    If .flags.Envenenado > 0 Then
                    
                        .flags.Envenenado = 0
                        
                        Call WriteConsoleMsg(UserIndex, "Te curaste del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
    
                        'Quitamos del inv el Item
                        Call QuitarInvItem(UserIndex, Slot)
                        
                        'Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible > 0 Then
                            Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                        End If
                    End If
                    
                Case 6  'Pocion Negra
                
                    If .flags.Privilegios And PlayerType.User Then
                        Call QuitarInvItem(UserIndex, Slot)
                        Call UserDie(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                    End If
           End Select
                 
        .flags.TomoPocion = True

         Case otBebida
         
            .Stats.MinSed = .Stats.MinSed + Obj.MinSed
            If .Stats.MinSed > 100 Then
                .Stats.MinSed = 100
            End If
            
            Call WriteUpdateHungerAndThirst(UserIndex)
            
            'Quitamos del inv el Item
            Call QuitarInvItem(UserIndex, Slot)
            
            'Los admin invisibles solo producen sonidos a si mismos
            If .flags.AdminInvisible > 0 Then
                Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToUserAreaButIndex, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
            End If
        
        Case otLlave
        
            If .flags.TargetObjIndex = 0 Then
                Exit Sub
            End If
            
            TargObj = ObjData(.flags.TargetObjIndex)
            '¿El objeto clickeado es una puerta?
            If TargObj.Type = otPuerta Then
                '¿Esta cerrada?
                If TargObj.Cerrada Then
                      '¿Cerrada con llave?
                      If TargObj.Llave > 0 Then
                         If TargObj.clave = Obj.clave Then
             
                            MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index _
                            = ObjData(MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index).IndexCerrada
                            .flags.TargetObjIndex = MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index
                            Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                         Else
                            Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                         End If
                      Else
                         If TargObj.clave = Obj.clave Then
                            MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index _
                            = ObjData(MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index).IndexCerradaLlave
                            Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                            .flags.TargetObjIndex = MapData(.flags.TargetObjX, .flags.TargetObjY).ObjInfo.index
                            Exit Sub
                         Else
                            Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                         End If
                      End If
                Else
                      Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                      Exit Sub
                End If
            End If
        
        Case otBotellaVacia
        
            If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                Call WriteConsoleMsg(UserIndex, "No hay agua ahí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            MiObj.Amount = 1
            MiObj.index = ObjData(.Inv.Obj(Slot).index).IndexAbierta
            
            Call QuitarInvItem(UserIndex, Slot)
            
            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
            End If
            
        Case otBotellaLlena
        
            .Stats.MinSed = .Stats.MinSed + Obj.MinSed
            
            If .Stats.MinSed > 100 Then
                .Stats.MinSed = 100
            End If
            
            Call WriteUpdateHungerAndThirst(UserIndex)
            
            MiObj.Amount = 1
            MiObj.index = ObjData(.Inv.Obj(Slot).index).IndexCerrada
            
            Call QuitarInvItem(UserIndex, Slot)
            
            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
            End If
        
        Case otPergamino
        
            If .Stats.MaxMan > 0 Then
                If .Stats.MinHam > 0 And .Stats.MinSed > 0 Then
                    Call AgregarHechizo(UserIndex, Slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y/o sediento.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No tenés conocimientos de las artes mágicas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        Case otPasaje
        
            If .Stats.Muerto Then
                Exit Sub
            End If
            
            If .flags.TargetNpcTipo <> Pirata Then
                Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Distancia(NpcList(.flags.TargetNpc).Pos, .Pos) > 4 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            If .Pos.Map <> Obj.DesdeMap Then
                Call WriteConsoleMsg(UserIndex, "¡El pasaje no lo compraste aquí! ¡Lárgate!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Not MapaValido(Obj.HastaMap) Then
                Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no está disponible. Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            If .Skills.Skill(Obj.NecesitasSkill).Elv < Obj.CantidadSkill Then
                Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, necesitás " & Obj.CantidadSkill & " puntos de habilidad en Navegación para que pueda llevarte.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
           
            Call WarpUserChar(UserIndex, Obj.HastaMap, Obj.HastaX, Obj.HastaY, True)
            
            Call WriteConsoleMsg(UserIndex, "Viajaste por varios días, te sentís exhausto.", FontTypeNames.FONTTYPE_INFO)
            
            .Stats.MinSed = 0
            .Stats.MinHam = 0
                        
            Call WriteUpdateHungerAndThirst(UserIndex)
            
            Call QuitarInvItem(UserIndex, Slot)
                        
        Case otInstrumento
        
            If GuildaPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.Map).PK = False Then
                    Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible > 0 Then
                    Call EnviarDatosASlot(UserIndex, Msg_SoundFX(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.toMap, .Pos.Map, Msg_SoundFX(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
                Exit Sub
            End If
            
            If .flags.AdminInvisible > 0 Then
                Call EnviarDatosASlot(UserIndex, Msg_SoundFX(Obj.Snd1, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(Obj.Snd1, .Pos.X, .Pos.Y))
            End If
           
        Case otBarco
            
            If .Stats.Muerto Then
                Exit Sub
            End If
                
            'Verifica si esta aproximado al agua antes de permitirle navegar
            If .Stats.Elv < 25 Then
                If .Clase <> eClass.Pirat Then
                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares tenés que ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
                    If .Stats.Elv < 20 Then
                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares tenés que ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
            End If
            
            If ((LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) _
                    Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) _
                    Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) _
                    Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) _
                    And Not .flags.Navegando) _
                    Or (.flags.Navegando And _
                    (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, False, True) _
                    Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, False, True) _
                    Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, False, True) _
                    Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, False, True))) Then
                Call DoNavega(UserIndex, Obj, Slot)
            Else
                If .flags.Navegando Then
                    Call WriteConsoleMsg(UserIndex, "Tenés que acercarte a tierra para salir del barco.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Tenés que acercarte al agua para usar el barco.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
    End Select
    
    End With

End Sub

Public Sub UseBeltInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Handels the usage of Items from inventory box.

    Dim Obj As ObjData
    Dim ObjIndex As Integer
    Dim MiObj As Obj
    
    With UserList(UserIndex)
        
        Obj = ObjData(.Belt.Obj(Slot).index)

        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then
            Exit Sub
        End If
        
        ObjIndex = .Belt.Obj(Slot).index
    
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Exit Sub
        End If
        
        .flags.TipoPocion = Obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = Obj.DuracionEfecto
    
                If .Stats.Atributos(eAtributos.Agilidad) < 2 * .Stats.AtributosBackUP(eAtributos.Agilidad) Then
                    'Usa el Item
                    .Stats.Atributos(eAtributos.Agilidad) = .Stats.Atributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                    
                    If .Stats.Atributos(eAtributos.Agilidad) > MaxAtributos Then
                        .Stats.Atributos(eAtributos.Agilidad) = MaxAtributos
                    End If
                    
                    If .Stats.Atributos(eAtributos.Agilidad) > 2 * .Stats.AtributosBackUP(eAtributos.Agilidad) Then
                        .Stats.Atributos(eAtributos.Agilidad) = 2 * .Stats.AtributosBackUP(eAtributos.Agilidad)
                    End If
                    
                    Call WriteUpdateDexterity(UserIndex)
                End If
    
                'Quitamos del cinturón el Item
                Call QuitarBeltItem(UserIndex, Slot)
                
                'Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible > 0 Then
                    Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                End If
    
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = Obj.DuracionEfecto
            
                If .Stats.Atributos(eAtributos.Fuerza) < 2 * .Stats.AtributosBackUP(Fuerza) Then
                    .Stats.Atributos(eAtributos.Fuerza) = .Stats.Atributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
    
                    If .Stats.Atributos(eAtributos.Fuerza) > MaxAtributos Then
                        .Stats.Atributos(eAtributos.Fuerza) = MaxAtributos
                    End If
                    
                    If .Stats.Atributos(eAtributos.Fuerza) > 2 * .Stats.AtributosBackUP(Fuerza) Then
                        .Stats.Atributos(eAtributos.Fuerza) = 2 * .Stats.AtributosBackUP(Fuerza)
                    End If
                    
                    Call WriteUpdateStrenght(UserIndex)
                End If
            
                'Quitamos del cinturón el Item
                Call QuitarBeltItem(UserIndex, Slot)
                
                'Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible > 0 Then
                    Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                End If
                
            Case 3 'Pocion roja, restaura HP
            
                'Usa el Item SI NO TIENE LA VIDA COMPLETA
                If .Stats.MinHP < .Stats.MaxHP Then
                    .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                    Call WriteUpdateHP(UserIndex)
                    
                    'Quitamos del cinturón el Item
                    Call QuitarBeltItem(UserIndex, Slot)
                    
                    'Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible > 0 Then
                        Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    End If
                End If
                
            Case 4 'Pocion azul, restaura MANA
            
                'Usa el Item SI NO TIENE LA MANÁ COMPLETA
                If .Stats.MinMan < .Stats.MaxMan Then
                    .Stats.MinMan = .Stats.MinMan + Porcentaje(.Stats.MaxMan, 5) + .Stats.Elv * 0.33 + 50 \ .Stats.Elv
                    Call WriteUpdateMana(UserIndex)
    
                    'Quitamos del cinturón el Item
                    Call QuitarBeltItem(UserIndex, Slot)
                    
                    'Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible > 0 Then
                        Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    End If
                End If
                
            Case 5 'Pocion violeta
            
                If .flags.Envenenado > 0 Then
                
                    .flags.Envenenado = 0
                    
                    Call WriteConsoleMsg(UserIndex, "Te curaste del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
    
                    'Quitamos del cinturón el Item
                    Call QuitarBeltItem(UserIndex, Slot)
                    
                    'Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible > 0 Then
                        Call EnviarDatosASlot(UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, Msg_SoundFX(SND_DRINK, .Pos.X, .Pos.Y))
                    End If
                End If
                
            Case 6  'Pocion Negra
            
                If .flags.Privilegios And PlayerType.User Then
                    Call QuitarBeltItem(UserIndex, Slot)
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
        End Select
             
        .flags.TomoPocion = True

    End With

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

    With ObjData(index)
        ItemSeCae = (.Guild <> 1 Or .NoSeCae = 0) And _
                    .Type <> otLlave And _
                    .Type <> otBarco And _
                    .NoSeCae = 0
    End With

End Function

Public Sub TirarItemsAlMorir(ByVal UserIndex As Integer)
    
    Dim i As Byte
    Dim ItemIndex As Integer
    Dim Calc As Long
    
        With UserList(UserIndex)
        For i = 1 To MaxInvSlots
            ItemIndex = .Inv.Obj(i).index
            
            If ItemIndex > 0 Then
                 If ItemSeCae(ItemIndex) Then
                 
                    Dim Probabilidad As Long 'Porcentaje
                    
                    Select Case ObjData(ItemIndex).Type
                        
                        Case otPocion, otFlecha
                            Probabilidad = 66

                        Case otLeña, otMineral
                            Probabilidad = 33
                            
                        Case otArma
                            Probabilidad = 25
                            
                            If ObjData(ItemIndex).Valor > 0 Then
                                Calc = CLng(.Stats.Elv) * 1000
                                Calc = ObjData(ItemIndex).Valor \ Calc
                                
                                If Calc > 1 Then
                                    Probabilidad = Probabilidad + Probabilidad * 0.05 * Calc
                                End If
                            End If
                                    
                        Case otEscudo, otCasco, otInstrumento
                            Probabilidad = 20
                            
                            If ObjData(ItemIndex).Valor > 0 Then
                                Calc = CLng(.Stats.Elv) * 1000
                                Calc = ObjData(ItemIndex).Valor \ Calc
                                
                                If Calc > 1 Then
                                    Probabilidad = Probabilidad + Probabilidad * 0.1 * Calc
                                End If
                            End If
                                    
                        Case otArmadura, otAnillo
                            Probabilidad = 15
                            
                            If ObjData(ItemIndex).Valor > 0 Then
                                Calc = CLng(.Stats.Elv) * 1000
                                Calc = ObjData(ItemIndex).Valor \ Calc
                                
                                If Calc > 1 Then
                                    Probabilidad = Probabilidad + Probabilidad * 0.15 * Calc
                                End If
                            End If
                            
                        Case otPasaje, otPergamino
                            Probabilidad = 10
                    
                    End Select
                    
                    If RandomNumber(0, 100) < Probabilidad Then
                        If .Inv.Obj(i).Amount > 0 Then
                            Call DropObj(UserIndex, i, RandomNumber(1, .Inv.Obj(i).Amount))
                        Else
                            Call DropObj(UserIndex, i, 1)
                        End If
                    End If
                End If
            End If
        Next i
    End With
End Sub
