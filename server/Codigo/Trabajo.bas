Attribute VB_Name = "Trabajo"
Option Explicit

Private Const GASTO_ENERGIA_TRABAJAR As Byte = 3

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'Chequea si ya debe mostrarse

On Error GoTo ErrHandler
    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        If .Counters.TiempoOculto < 1 Then
            
            If .Clase = eClass.Bandit Then
                .Counters.TiempoOculto = Int(IntervaloOculto * 0.5)
            Else
                .Counters.TiempoOculto = IntervaloOculto
            End If
            
            If .Clase = eClass.Hunter And .Skills.Skill(eSkill.Ocultarse).Elv > 90 Then
                If .Inv.Body = 648 Or .Inv.Body = 360 Then
                    Exit Sub
                End If
            End If
            
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando Then
                If .Clase = eClass.Pirat Then
                    'Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.Invisible < 1 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en PUBLIC SUB DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

    Dim Suerte As Double
    Dim res As Integer
    Dim Skill As Integer
    
    With UserList(UserIndex)
        Skill = .Skills.Skill(eSkill.Ocultarse).Elv
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            .Counters.TiempoOculto = Suerte
            
            'No es pirata o es uno sin barca
            If Not .flags.Navegando Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        
            'Es un pirata navegando
            Else
                'Le cambiamos el body a galeon fantasmal
                .Char.Body = iFragataFantasmal
                'Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                    NingunEscudo, NingunCasco)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)
        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en PUBLIC SUB DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Byte)

    Dim ModNave As Single
    
    With UserList(UserIndex)
        ModNave = ModNavegacion(.Clase, UserIndex)
        
        If .Skills.Skill(eSkill.Navegacion).Elv / ModNave < Barco.MinSkill Then
            If Barco.MinSkill * ModNave > MaxSkillPoints Then
                Call WriteConsoleMsg(UserIndex, "Solo los piratas pueden manejar este barco.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Para navegar este barco necesitás " & Barco.MinSkill * ModNave & " puntos de habilidad en Navegación.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        .Inv.Ship = .Inv.Obj(Slot).index
        
        'No estaba navegando
        If Not .flags.Navegando Then
            Call ToogleBoatBody(UserIndex)
            
            If .Clase = eClass.Pirat Then
                If .flags.Oculto > 0 Then
                    .flags.Oculto = 0
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
            
            'Comienza a navegar
            .flags.Navegando = True
            
        'Estaba navegando
        Else
            If .Clase = eClass.Pirat Then
                If .flags.Oculto > 0 Then
                    'Al desequipar barca, perdió el ocultar
                    .flags.Oculto = 0
                    .Counters.Ocultando = 0
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
                
            .Char.Head = .OrigChar.Head
                
            If .Inv.Head > 0 Then
                .Char.HeadAnim = ObjData(.Inv.Head).HeadAnim
            Else
                .Char.HeadAnim = NingunCasco
            End If
        
            If .Inv.Body > 0 Then
                .Char.Body = ObjData(.Inv.Body).BodyAnim
                .flags.Desnudo = False
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
                
            .Char.WeaponAnim = GetWeaponAnim(UserIndex)
            
            If UsaEscudo(UserIndex) > 0 Then
                .Char.ShieldAnim = ObjData(.Inv.LeftHand).ShieldAnim
            Else
                .Char.ShieldAnim = NingunEscudo
            End If
                    
            'Termina de navegar
            .flags.Navegando = False
        End If
        
        Call WriteInventorySlot(UserIndex, Slot)
        
        'Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim)
    End With
    
    Call WriteNavigateToggle(UserIndex)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).MinSkill <= .Skills.Skill(eSkill.Mineria).Elv Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Para trabajar este mineral necesitás " & ObjData(.flags.TargetObjInvIndex).MinSkill & " puntos de habilidad en Minería.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & ": " & Err.description)

End Sub

Public Sub FundirArmas(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).Type = otArma Then
                If ObjData(.flags.TargetObjInvIndex).SkHerreria <= .Skills.Skill(eSkill.Herreria).Elv Then
                    Call DoFundir(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "Para fundir este objeto necesitás " & ObjData(.flags.TargetObjInvIndex).SkHerreria & " puntos de habilidad en Herrería.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en FundirArmas. Error " & Err.Number & ": " & Err.description)
End Sub

Public Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean

    Dim i As Integer
    Dim Total As Long
    For i = 1 To MaxInvSlots
        If UserList(UserIndex).Inv.Obj(i).index = ItemIndex Then
            Total = Total + UserList(UserIndex).Inv.Obj(i).Amount
        End If
    Next i
    
    If Cant <= Total Then
        TieneObjetos = True
        Exit Function
    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer)

    Dim i As Integer
    
    For i = 1 To MaxInvSlots
        With UserList(UserIndex).Inv.Obj(i)
            If .index = ItemIndex Then
                Call QuitarInvItem(UserIndex, i, Cant)
            End If
        End With
    Next i

End Sub

Public Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
    With ObjData(ItemIndex)
        If .LingH > 0 Then
            Call QuitarObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex)
        End If
            
        If .LingP > 0 Then
            Call QuitarObjetos(LingotePlata, .LingP * CantidadItems, UserIndex)
        End If
        
        If .LingO > 0 Then
            Call QuitarObjetos(LingoteOro, .LingO * CantidadItems, UserIndex)
        End If
    End With
End Sub

Public Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
    With ObjData(ItemIndex)
        If .Madera > 0 Then
            Call QuitarObjetos(Leña, .Madera * CantidadItems, UserIndex)
        End If
        
        If .MaderaElfica > 0 Then
            Call QuitarObjetos(LeñaElfica, .MaderaElfica * CantidadItems, UserIndex)
        End If
    End With
End Sub

Public Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer, Optional ByVal ShowMsg As Boolean = False) As Boolean
    With ObjData(ItemIndex)
        If .Madera > 0 Then
            
            Dim Madera As Long
            
            Madera = Cantidad * .Madera
        
            If Not TieneObjetos(Leña, Madera, UserIndex) Then
                If ShowMsg Then
                    Call WriteConsoleMsg(UserIndex, "Para construir " & Cantidad & " " & ObjData(ItemIndex).name & " necesitás " & Madera & " leños.", FontTypeNames.FONTTYPE_INFO)
                End If
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            
            Madera = .MaderaElfica * Cantidad
            
            If Not TieneObjetos(LeñaElfica, Madera, UserIndex) Then
                If ShowMsg Then
                    Call WriteConsoleMsg(UserIndex, "Para construir " & Cantidad & " " & ObjData(ItemIndex).name & " necesitás " & Madera & " leños élficos.", FontTypeNames.FONTTYPE_INFO)
                End If
                
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
    
    End With
    CarpinteroTieneMateriales = True

End Function
 
Public Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
    With ObjData(ItemIndex)
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, .LingP * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, .LingO * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
    End With
    HerreroTieneMateriales = True
End Function

Public Function TieneMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .Madera > 0 Then
            If Not TieneObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
    End With
    
    TieneMaterialesUpgrade = True
End Function

Public Sub QuitarMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then
            Call QuitarObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        End If
        
        If .LingP > 0 Then
            Call QuitarObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        End If
        
        If .LingO > 0 Then
            Call QuitarObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        End If
        
        If .Madera > 0 Then
            Call QuitarObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        End If
        
        If .MaderaElfica > 0 Then
            Call QuitarObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        End If
    End With
    
    Call QuitarObjetos(ItemIndex, 1, UserIndex)
End Sub

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, CantidadItems) And UserList(UserIndex).Skills.Skill(eSkill.Herreria).Elv >= ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ArmasHerrero)
        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next i
    For i = 1 To UBound(ArmadurasHerrero)
        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function
        End If
    Next i
    PuedeConstruirHerreria = False
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    
    Dim CantidadItems As Integer
    Dim TieneMateriales As Boolean
    
    With UserList(UserIndex)
        CantidadItems = .Construir.PorCiclo
        
        If .Construir.Cantidad < CantidadItems Then _
            CantidadItems = .Construir.Cantidad
            
        If .Construir.Cantidad > 0 Then _
            .Construir.Cantidad = .Construir.Cantidad - CantidadItems
            
        If CantidadItems = 0 Then
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        If PuedeConstruirHerreria(ItemIndex) Then
            
            While CantidadItems > 0 And Not TieneMateriales
                If PuedeConstruir(UserIndex, ItemIndex, CantidadItems) Then
                    TieneMateriales = True
                Else
                    CantidadItems = CantidadItems - 1
                End If
            Wend
            
            'Chequeo si puede hacer al menos 1 Item
            If Not TieneMateriales Then
                Call WriteConsoleMsg(UserIndex, "No tenés suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
                Call WriteStopWorking(UserIndex)
                Exit Sub
            End If
            
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJAR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJAR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call HerreroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
            
            Dim MiObj As Obj
            
            MiObj.Amount = CantidadItems
            MiObj.index = ItemIndex
            
            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
            End If
            
            'Log de construcción de Items. Pablo (ToxicWaste) 10\09\07
            If ObjData(MiObj.index).Log = 1 Then
                Call LogDesarrollo(.name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.index).name)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Herreria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.x, .Pos.y))
        
            .Counters.Trabajando = .Counters.Trabajando + 1
        End If
    End With
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
    
    Dim i As Long
    
    For i = 1 To UBound(ObjCarpintero)
        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function
        End If
    Next i
    PuedeConstruirCarpintero = False
    
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    
    Dim CantidadItems As Integer
    Dim TieneMateriales As Boolean
    Dim WeaponIndex As Integer
    Dim OtroUserIndex As Integer
    
    With UserList(UserIndex)
        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxPoblacion Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
        
        WeaponIndex = .Inv.RightHand
    
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        CantidadItems = .Construir.PorCiclo
        
        If .Construir.Cantidad < CantidadItems Then _
            CantidadItems = .Construir.Cantidad
            
        If .Construir.Cantidad > 0 Then _
            .Construir.Cantidad = .Construir.Cantidad - CantidadItems
            
        If CantidadItems = 0 Then
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
    
        If PuedeConstruirCarpintero(ItemIndex) Then
            
            If .Skills.Skill(eSkill.Carpinteria).Elv < ObjData(ItemIndex).SkCarpinteria Then
                Exit Sub
            End If
           
            'Calculo cuantos item puede construir
            While CantidadItems > 0 And Not TieneMateriales
                If CarpinteroTieneMateriales(UserIndex, ItemIndex, CantidadItems) Then
                    TieneMateriales = True
                Else
                    CantidadItems = CantidadItems - 1
                End If
            Wend
            
            'No tiene los materiales ni para construir 1 item?
            If Not TieneMateriales Then
                ' Para que muestre el mensaje
                Call CarpinteroTieneMateriales(UserIndex, ItemIndex, 1, True)
                Call WriteStopWorking(UserIndex)
                Exit Sub
            End If
           
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJAR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJAR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
             
            Dim MiObj As Obj
            MiObj.Amount = CantidadItems
            MiObj.index = ItemIndex
            
            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
            End If
            
            'Log de construcción de Items. Pablo (ToxicWaste) 10\09\07
            If ObjData(MiObj.index).Log = 1 Then
                Call LogDesarrollo(.name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.index).name)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.x, .Pos.y))
        
            .Counters.Trabajando = .Counters.Trabajando + 1
        End If
    End With
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 3
        Case iMinerales.PlataCruda
            MineralesParaLingote = 5
        Case iMinerales.OroCrudo
            MineralesParaLingote = 10
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)

'Call LogTarea("PUBLIC SUB DoLingotes")
    Dim Slot As Byte
    Dim obji As Integer
    Dim CantidadItems As Integer
    Dim TieneMinerales As Boolean

    With UserList(UserIndex)
        CantidadItems = MaximoInt(1, .Stats.Elv)

        Slot = .flags.TargetObjInvSlot
        obji = .Inv.Obj(Slot).index
        
        While CantidadItems > 0 And Not TieneMinerales
            If .Inv.Obj(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        If ObjData(obji).Type <> otMineral Then
            Exit Sub
        End If
        
        If Not TieneMinerales Then
            Call WriteConsoleMsg(UserIndex, "Para hacer " & ObjData(ObjData(.flags.TargetObjInvIndex).LingoteIndex).name & " necesitás " & MineralesParaLingote(obji) & " " & ObjData(obji).name & ".", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call QuitarInvItem(UserIndex, Slot, MineralesParaLingote(obji) * CantidadItems)
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.index = ObjData(.flags.TargetObjInvIndex).LingoteIndex
        
        If Not MeterEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
        End If
        
        If Not .flags.UltimoMensaje = 11 Then
            Call WriteConsoleMsg(UserIndex, "¡Obtuviste " & CantidadItems & " lingote" & _
                            IIf(CantidadItems = 1, vbNullString, "s") & "!", FontTypeNames.FONTTYPE_INFO)
            .flags.UltimoMensaje = 11
        End If
            
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)
    
    Dim i As Integer
    Dim Num As Integer
    Dim Slot As Byte
    Dim Lingotes(2) As Integer
    
        With UserList(UserIndex)
            Slot = .flags.TargetObjInvSlot
            
            With .Inv.Obj(Slot)
                .Amount = .Amount - 1
                
                If .Amount < 1 Then
                    .Amount = 0
                    .index = 0
                End If
                
                Call WriteInventorySlot(UserIndex, Slot)
            End With
            
            Num = RandomNumber(20, 40)
            
            Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
            Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
            Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
        
        Dim MiObj(2) As Obj
        
        For i = 0 To 2
            MiObj(i).Amount = Lingotes(i)
            MiObj(i).index = LingoteHierro + i 'Una gran negrada pero práctica
            If MiObj(i).Amount > 0 Then
                If Not MeterEnInventario(UserIndex, MiObj(i)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(i), , UserIndex)
                End If
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido el " & Num & "% de los lingotes utilizados para la construcción del objeto!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    
    End With
    
End Sub

Public Sub DoUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With UserList(UserIndex)
        'Sacamos energía
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA_TRABAJAR Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJAR
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No tenés suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If ItemUpgrade < 1 Then
            Exit Sub
        End If
        
        If Not TieneMaterialesUpgrade(UserIndex, ItemIndex) Then
            Exit Sub
        End If
        
        If PuedeConstruirHerreria(ItemUpgrade) Then
        
            If .Inv.RightHand <> MARTILLO_HERRERO Then
                Call WriteConsoleMsg(UserIndex, "Debes equiparte el martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .Skills.Skill(eSkill.Herreria).Elv < ObjData(ItemUpgrade).SkHerreria Then
                Call WriteConsoleMsg(UserIndex, "Para mejorar este objeto necesitás " & ObjData(ItemUpgrade).SkHerreria & " puntos de habilidad en Herrería.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Select Case ObjData(ItemIndex).Type
                Case otArma
                    Call WriteConsoleMsg(UserIndex, "Has mejorado el arma.", FontTypeNames.FONTTYPE_INFO)
                    
                Case otEscudo 'Todavía no hay, pero just in case
                    Call WriteConsoleMsg(UserIndex, "Has mejorado el escudo.", FontTypeNames.FONTTYPE_INFO)
                
                Case otCasco
                    Call WriteConsoleMsg(UserIndex, "Has mejorado el casco.", FontTypeNames.FONTTYPE_INFO)
                
                Case otArmadura
                    Call WriteConsoleMsg(UserIndex, "Has mejorado la armadura.", FontTypeNames.FONTTYPE_INFO)
            End Select
            
            Call SubirSkill(UserIndex, eSkill.Herreria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.x, .Pos.y))
        
        ElseIf PuedeConstruirCarpintero(ItemUpgrade) Then
        
            If .Inv.RightHand <> SERRUCHO_CARPINTERO Then
                Call WriteConsoleMsg(UserIndex, "Debes equiparte el serrucho.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .Skills.Skill(eSkill.Carpinteria).Elv < ObjData(ItemUpgrade).SkCarpinteria Then
                Call WriteConsoleMsg(UserIndex, "Para mejorar este objeto necesitás " & ObjData(ItemUpgrade).SkCarpinteria & " puntos de habilidad en Carpintería.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Select Case ObjData(ItemIndex).Type
                Case otFlecha
                    Call WriteConsoleMsg(UserIndex, "Has mejorado la flecha.", FontTypeNames.FONTTYPE_INFO)
                    
                Case otArma
                    Call WriteConsoleMsg(UserIndex, "Has mejorado el arma.", FontTypeNames.FONTTYPE_INFO)
                    
                Case otBarco
                    Call WriteConsoleMsg(UserIndex, "Has mejorado el barco.", FontTypeNames.FONTTYPE_INFO)
            End Select
            
            Call SubirSkill(UserIndex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.x, .Pos.y))
        Else
            Exit Sub
        End If
        
        Call QuitarMaterialesUpgrade(UserIndex, ItemIndex)
        
        Dim MiObj As Obj
        MiObj.Amount = 1
        MiObj.index = ItemUpgrade
        
        If Not MeterEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
        End If
        
        If ObjData(ItemIndex).Log Then
            Call LogDesarrollo(.name & " ha mejorado el ítem " & ObjData(ItemIndex).name & " a " & ObjData(ItemUpgrade).name)
        End If
                
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
End Sub

Public Function ModNavegacion(ByVal Clase As eClass, ByVal UserIndex As Integer) As Single
    If Clase = eClass.Pirat Then
        ModNavegacion = 1
    Else
        ModNavegacion = 2
    End If
End Function

Public Function FreeMascotaIndex(ByVal UserIndex As Integer) As Byte
'Busca un indice libre de Mascotas, revisando los types y no los indices de los npcs

    Dim j As Integer
    For j = 1 To MaxPets
        If UserList(UserIndex).Pets.Pet(j).Tipo < 1 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Public Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

On Error GoTo ErrHandler

    Dim puntosDomar As Integer
    Dim puntosRequeridos As Integer
    
    If NpcList(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    With UserList(UserIndex)
        If .Pets.NroALaVez < MaxPetsALaVez Then
            
            If NpcList(NpcIndex).MaestroNpc > 0 Or NpcList(NpcIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            puntosDomar = CInt(.Stats.Atributos(eAtributos.Carisma)) * CInt(.Skills.Skill(eSkill.Domar).Elv)
            
            Select Case .Clase
                Case eClass.Druid
                    If .Inv.Ring = FLAUTAELFICA Then
                        puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.7
                    ElseIf .Inv.Ring = FLAUTAMAGICA Then
                        puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.85
                    End If
                    
                Case eClass.Bard
                    If .Inv.Ring = LAUDELFICO Then
                        puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.8
                    ElseIf .Inv.Ring = LAUDMAGICO Then
                        puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.9
                    End If
                    
                Case eClass.Hunter
                    puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.9
                    
                Case eClass.Cleric
                    If ObjData(.Inv.Ring).Type = otInstrumento Then
                        puntosRequeridos = NpcList(NpcIndex).flags.Domable * 0.9
                    End If
                    
                Case Else
                    puntosRequeridos = NpcList(NpcIndex).flags.Domable
            End Select
            
            If puntosRequeridos > -1 Then 'RandomNumber(NpcList(NpcIndex).Stats.MinHP, NpcList(NpcIndex).Stats.MaxHP) < NpcList(NpcIndex).Stats.MaxHP * (puntosDomar \ puntosRequeridos) \ 10
                Dim index As Byte
            
                index = FreeMascotaIndex(UserIndex)
                
                .Pets.Pet(index).Nombre = NpcList(NpcIndex).name
                .Pets.Pet(index).index = NpcIndex
                .Pets.Pet(index).Tipo = NpcList(NpcIndex).Numero
                
                .Pets.Pet(index).Lvl = NpcList(NpcIndex).Lvl
                
                .Pets.Pet(index).MinHP = NpcList(NpcIndex).Stats.MinHP
                .Pets.Pet(index).MaxHP = NpcList(NpcIndex).Stats.MaxHP
                
                .Pets.Pet(index).MinHit = NpcList(NpcIndex).Stats.MinHit
                .Pets.Pet(index).MaxHit = NpcList(NpcIndex).Stats.MaxHit
                
                .Pets.Pet(index).Def = NpcList(NpcIndex).Stats.Def
                .Pets.Pet(index).DefM = NpcList(NpcIndex).Stats.DefM
                
                NpcList(NpcIndex).MaestroUser = UserIndex
                
                .Pets.Nro = .Pets.Nro + 1
                .Pets.NroALaVez = .Pets.NroALaVez + 1
                
                Call FollowAmo(NpcIndex)
                
                If NpcList(NpcIndex).flags.Respawn > 0 Then
                    Call CrearNpc(NpcList(NpcIndex).Numero, NpcList(NpcIndex).Pos.map, NpcList(NpcIndex).Orig)
                End If
                
                Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).name & " (Nv." & NpcList(NpcIndex).Lvl - 1 & ") es ahora tu mascota.", FontTypeNames.FONTTYPE_INFO)

                Call SubirSkill(UserIndex, eSkill.Domar, True)

            Else
                Call WriteConsoleMsg(UserIndex, "No has podido domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(UserIndex, eSkill.Domar, False)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No podés controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & ": " & Err.description)

End Sub

Public Sub DoAdminInvisible(ByVal UserIndex As Integer)
'Makes an admin invisible o visible.

    With UserList(UserIndex)
        If .flags.AdminInvisible < 1 Then
            'Sacamos el mimetizmo
            If .flags.Mimetizado Then
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.HeadAnim = .CharMimetizado.HeadAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = False
                'Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False
            End If
            
            .flags.AdminInvisible = 1
            .flags.Invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.Body
            .flags.OldHead = .Char.Head
            .Char.Body = 0
            .Char.Head = 0
            
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharRemove(.Char.CharIndex))
            
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            
        Else
            .flags.AdminInvisible = 0
            .flags.Invisible = 0
            
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.Body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            'Solo el admin sabe que se hace visible
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharChange(.Char.Body, .Char.Head, .Char.Heading, _
            .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.HeadAnim))
            
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
             
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.x, .Pos.y)
        End If
    End With
    
End Sub

Public Sub TratarDeHacerFogata(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)

    Dim Suerte As Byte
    Dim exito As Byte
    Dim Obj As Obj
    Dim posMadera As WorldPos
    
    If Not LegalPos(map, x, y) Then
        Exit Sub
    End If
    
    With posMadera
        .map = map
        .x = x
        .y = y
    End With
    
    If maps(map).mapData(x, y).ObjInfo.index <> 58 Then
        Call WriteConsoleMsg(UserIndex, "Necesitás hacer click sobre leña para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.Muerto Then
        Call WriteConsoleMsg(UserIndex, "No podés hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If maps(map).mapData(x, y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(UserIndex, "Necesitás por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim SupervivenciaSkill As Byte
    
    SupervivenciaSkill = UserList(UserIndex).Skills.Skill(eSkill.Supervivencia).Elv
    
    If SupervivenciaSkill > 1 And SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill >= 6 And SupervivenciaSkill <= 34 Then
        Suerte = 2
    ElseIf SupervivenciaSkill >= 35 Then
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    If exito = 1 Then
        Obj.index = FOGATA_APAG
        Obj.Amount = maps(map).mapData(x, y).ObjInfo.Amount \ 3
        
        Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, map, x, y)
        
        'Seteamos la fogata como el nuevo TargetObjIndex del user
        UserList(UserIndex).flags.TargetObjIndex = FOGATA_APAG
        
        Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 10
        End If
        
        Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
    End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)

On Error GoTo ErrHandler
    
    Dim Suerte As Integer
    Dim res As Integer
    Dim CantidadItems As Integer
    
    Call QuitarSta(UserIndex, EsfuerzoPescar)
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Skills.Skill(eSkill.Pesca).Elv
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 45)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 5 Then
        Dim MiObj As Obj
        
        MiObj.index = Pescado
        
        MiObj.Amount = RandomNumber(UserList(UserIndex).Stats.Elv * 0.8, UserList(UserIndex).Stats.Elv * 1.2)
        
        If Not MeterEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, , UserIndex)
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has pescado " & MiObj.Amount & " peces.", FontTypeNames.FONTTYPE_INFO)
        
        Call SubirSkill(UserIndex, eSkill.Pesca, True)
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
            Call WriteConsoleMsg(UserIndex, "No has pescado nada.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 6
        End If
        
        Call SubirSkill(UserIndex, eSkill.Pesca, False)
    End If
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & ": " & Err.description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim iSkill As Integer
    Dim Suerte As Integer
    Dim res As Integer
    
    Call QuitarSta(UserIndex, EsfuerzoPescar)

    iSkill = UserList(UserIndex).Skills.Skill(eSkill.Pesca).Elv
    
    'm = (60-11)\(1-10)
    'y = mx - m*10 + 11
    
    Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 45)
    
    If Suerte > 0 Then
        res = RandomNumber(1, Suerte)
        
        If res < 6 Then
            Dim MiObj As Obj
            Dim PecesPosibles(1 To 4) As Integer
            
            PecesPosibles(1) = PESCADO1
            PecesPosibles(2) = PESCADO2
            PecesPosibles(3) = PESCADO3
            PecesPosibles(4) = PESCADO4
            
            MiObj.Amount = RandomNumber(UserList(UserIndex).Stats.Elv * 0.8, UserList(UserIndex).Stats.Elv * 1.2)
                        
            MiObj.index = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
            
            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, , UserIndex)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Pesca, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Pesca, False)
        End If
    End If
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

On Error GoTo ErrHandler

    If Not MapInfo(UserList(VictimaIndex).Pos.map).PK Then
        Exit Sub
    End If
        
    If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then
        Exit Sub
    End If
    
    With UserList(LadrOnIndex)
        
        'Tiene energia?
        If .Stats.MinSta < 15 Then
            Call WriteConsoleMsg(LadrOnIndex, "No tenés suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Inv.Ring = GUANTE_HURTO Then
            GuantesHurto = True
        End If
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte As Integer
            Dim res As Integer
            Dim RobarSkill As Byte
            
            RobarSkill = .Skills.Skill(eSkill.Robar).Elv
                
            If RobarSkill < 20 And RobarSkill >= -1 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 And RobarSkill >= 11 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 And RobarSkill >= 21 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 And RobarSkill >= 31 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 And RobarSkill >= 41 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 And RobarSkill >= 51 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 And RobarSkill >= 61 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 And RobarSkill >= 71 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 And RobarSkill >= 81 Then
                Suerte = 10
            ElseIf RobarSkill < 100 And RobarSkill >= 91 Then
                Suerte = 7
            ElseIf RobarSkill = 100 Then
                Suerte = 5
            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
               
                If (RandomNumber(1, 50) < 25) And (.Clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else 'Roba oro
                    If UserList(VictimaIndex).Stats.Gld > 0 Then
                        Dim N As Integer
                        
                        If .Clase = eClass.Thief Then
                        'Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                N = RandomNumber(.Stats.Elv * 50, .Stats.Elv * 100)
                            Else
                                N = RandomNumber(.Stats.Elv * 25, .Stats.Elv * 50)
                            End If
                        Else
                            N = RandomNumber(1, 100)
                        End If
                        
                        N = N * MultiplicadorGld
                        
                        If N > UserList(VictimaIndex).Stats.Gld Then
                            N = UserList(VictimaIndex).Stats.Gld
                        End If
                        
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - N
                        
                        .Stats.Gld = .Stats.Gld + N
                        If .Stats.Gld > MaxOro Then
                            .Stats.Gld = MaxOro
                        End If
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Call FlushBuffer(VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(VictimaIndex)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)
            End If
        End If
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & ": " & Err.description)

End Sub

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Byte) As Boolean
'Check if one Item is stealable

    Dim OI As Integer
    
    OI = UserList(VictimaIndex).Inv.Obj(Slot).index
    
    ObjEsRobable = ItemSeCae(OI)

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Try to steal an Item to another Char

    Dim flag As Boolean
    Dim i As Integer
    flag = False
    
    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
        i = 1
        Do While Not flag And i <= MaxInvSlots
            'Hay objeto en este slot?
            If UserList(VictimaIndex).Inv.Obj(i).index > 0 Then
                If ObjEsRobable(VictimaIndex, i) Then
                    If RandomNumber(1, 10) < 4 Then
                        flag = True
                    End If
                End If
            End If
            If Not flag Then
                i = i + 1
            End If
        Loop
    Else
        i = 20
        Do While Not flag And i > 0
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Inv.Obj(i).index > 0 Then
            If ObjEsRobable(VictimaIndex, i) Then
                If RandomNumber(1, 10) < 4 Then
                    flag = True
                End If
            End If
        End If
        If Not flag Then
            i = i - 1
        End If
        Loop
    End If
    
    If flag Then
        Dim MiObj As Obj
        Dim Num As Byte
        Dim ObjAmount As Integer
        
        ObjAmount = UserList(VictimaIndex).Inv.Obj(i).Amount
        
        'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
        Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
        MiObj.Amount = Num
        MiObj.index = UserList(VictimaIndex).Inv.Obj(i).index
        
        UserList(VictimaIndex).Inv.Obj(i).Amount = ObjAmount - Num
                    
        If UserList(VictimaIndex).Inv.Obj(i).Amount < 1 Then
            Call QuitarInvItem(VictimaIndex, i)
        End If
                                    
        If Not MeterEnInventario(LadrOnIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj, , LadrOnIndex)
        End If
        
        If UserList(LadrOnIndex).Clase = eClass.Thief Then
            Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.index).name, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.index).name, FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    'If exiting, cancel de quien es robado
    Call CancelExit(VictimaIndex)

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)

On Error GoTo ErrHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then
        UserList(UserIndex).Stats.MinSta = 0
    End If
    
    Call WriteUpdateSta(UserIndex)
    
Exit Sub

ErrHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & ": " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, Optional ByVal DarMaderaElfica As Boolean = False)

On Error GoTo ErrHandler

    Dim Suerte As Integer
    Dim res As Integer
    Dim CantidadItems As Integer
    
    Call QuitarSta(UserIndex, EsfuerzoTalar)
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Skills.Skill(eSkill.Talar).Elv
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 45)
    
    res = RandomNumber(0, Suerte)
    
    If res <= 7 Then
        Dim MiObj As Obj
                
        MiObj.Amount = Round(RandomNumber(UserList(UserIndex).Stats.Elv * 0.8, UserList(UserIndex).Stats.Elv * 1.2) / 10)

        MiObj.index = IIf(DarMaderaElfica, LeñaElfica, Leña)
        
        If Not MeterEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, , UserIndex)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Talar, True)
    Else
        Call SubirSkill(UserIndex, eSkill.Talar, False)
    End If
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMinar(ByVal UserIndex As Integer)

On Error GoTo ErrHandler

    Dim Suerte As Integer
    Dim res As Integer
    Dim CantidadItems As Integer
    
    With UserList(UserIndex)
    
        If .flags.TargetObjIndex < 1 Then
            Exit Sub
        End If

        Call QuitarSta(UserIndex, EsfuerzoExcavar)
        
        Dim Skill As Integer
        Skill = .Skills.Skill(eSkill.Mineria).Elv
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 45)
        
        res = RandomNumber(0, Suerte)
        
        Dim Modificador As Byte
        
        Select Case ObjData(.flags.TargetObjIndex).MineralIndex
            Case iMinerales.HierroCrudo
                Modificador = 10
            Case iMinerales.PlataCruda
                Modificador = 7
            Case iMinerales.OroCrudo
                Modificador = 5
        End Select
        
        If res <= Modificador Then
            Dim MiObj As Obj
            
            MiObj.index = ObjData(.flags.TargetObjIndex).MineralIndex
            
            MiObj.Amount = Round(RandomNumber(.Stats.Elv * 0.8, .Stats.Elv * 1.2) / 10)

            If Not MeterEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj, , UserIndex)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Mineria, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Mineria, False)
        End If
        
        .Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en PUBLIC SUB DoMinar")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte As Integer
        Dim res As Integer
        Dim Cant As Integer
        Dim MeditarSkill As Byte
        
        'Esperamos a que se termine de concentrar
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub
        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True
        End If
            
        If .Stats.MinMan >= .Stats.MaxMan Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCreateCharFX(.Char.CharIndex))
            Exit Sub
        End If
        
        MeditarSkill = .Skills.Skill(eSkill.Meditar).Elv
        
        If MeditarSkill < 20 And MeditarSkill >= -1 Then
            Suerte = 13
        ElseIf MeditarSkill <= 20 And MeditarSkill >= 11 Then
            Suerte = 12
        ElseIf MeditarSkill <= 30 And MeditarSkill >= 21 Then
            Suerte = 11
        ElseIf MeditarSkill <= 40 And MeditarSkill >= 31 Then
            Suerte = 10
        ElseIf MeditarSkill <= 50 And MeditarSkill >= 41 Then
            Suerte = 9
        ElseIf MeditarSkill <= 60 And MeditarSkill >= 51 Then
            Suerte = 8
        ElseIf MeditarSkill <= 70 And MeditarSkill >= 61 Then
            Suerte = 7
        ElseIf MeditarSkill <= 80 And MeditarSkill >= 71 Then
            Suerte = 6
        ElseIf MeditarSkill <= 90 And MeditarSkill >= 81 Then
            Suerte = 5
        ElseIf MeditarSkill < 100 And MeditarSkill >= 91 Then
            Suerte = 4
        ElseIf MeditarSkill = 100 Then
            Suerte = 3
        End If
        
        res = RandomNumber(0, Suerte)
        
        If res < 10 Then
            Cant = Porcentaje(.Stats.MaxMan, 1)
            If Cant < 1 Then
                Cant = 1
            End If
            .Stats.MinMan = .Stats.MinMan + Cant
            
            Call WriteUpdateMana(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Meditar, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Meditar, False)
        End If
    End With

End Sub
Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'Unequips either shield, weapon or helmet from target user.

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado As Boolean
    
    With UserList(UserIndex)
        'Si no tiene guantes de hurto no desequipa.
        If .Inv.Ring <> GUANTE_HURTO Then
            Exit Sub
        End If
        
        'Si no esta solo con manos, no desequipa tampoco.
        If UsaArco(UserIndex) > 0 Or UsaArmaNoArco(UserIndex) > 0 Then
            Exit Sub
        End If
        
        WrestlingSkill = .Skills.Skill(eSkill.Wrestling).Elv
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.Elv * 0.66
   End With
   
   With UserList(VictimIndex)
        'Si tiene escudo, intenta desequiparlo
        If UsaEscudo(VictimIndex) > 0 Then
        
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                'Se lo desequipo
                Call Desequipar(VictimIndex, otEscudo)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.Elv < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
            
        ElseIf UsaArco(UserIndex) > 0 Or UsaArmaNoArco(UserIndex) > 0 Then

            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                'Se lo desequipo
                Call Desequipar(VictimIndex, otArma)
                
                Call WriteConsoleMsg(UserIndex, "¡Lograste desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.Elv < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente desarmó!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        'No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Inv.Head > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                'Se lo desequipo
                Call Desequipar(VictimIndex, otCasco)
                
                Call WriteConsoleMsg(UserIndex, "Lograste desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.Elv < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te desequipó el casco!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado Items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún Item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'Implements the pick pocket skill of the Bandit :)
    
    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then
        Exit Sub
    End If
    
    If UserList(UserIndex).Clase <> eClass.Bandit Then
        Exit Sub
    End If
    
    'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
    'Uso el slot de los anillos para "equipar" los guantes.
    'Y los reconozco porque les puse MinDefM y Max = 0
    If UserList(UserIndex).Inv.Ring <> GUANTE_HURTO Then
        Exit Sub
    End If
    
    Dim res As Integer
    res = RandomNumber(1, 100)
    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
            Call RobarObjeto(UserIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'Implements the special Skill of the Thief

    If UserList(VictimaIndex).flags.Paralizado > 0 Then
        Exit Sub
    End If
    
    If UserList(UserIndex).Clase <> eClass.Thief Then
        Exit Sub
    End If
        
    If UserList(UserIndex).Inv.Ring <> GUANTE_HURTO Then
        Exit Sub
    End If
        
    If RandomNumber(0, 100) < (UserList(UserIndex).Skills.Skill(eSkill.Wrestling).Elv \ 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado * 0.5
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageSetParalized(UserList(VictimaIndex).Char.CharIndex, 1))
        Call WritePosUpdate(VictimaIndex)
        Call WriteConsoleMsg(UserIndex, "Tu golpe paralizó a tu oponente.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡El golpe te dejó paralizado!", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Skills.Skill(eSkill.Wrestling).Elv
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.Elv * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, otArma)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(VictimIndex).Stats.Elv < 20 Then
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente desarmó!", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call FlushBuffer(VictimIndex)
        End If
    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
    MaxItemsConstruibles = MaximoInt(1, CInt((UserList(UserIndex).Stats.Elv - 4) \ 5))
End Function
