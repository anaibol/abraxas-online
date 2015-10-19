Attribute VB_Name = "Mod_Subastas"
Public Type c_subasta
    Actual As Boolean 'Sabemos si hay una subasta
    
    UserIndex As Integer 'UserIndex del Usuario Subastando
    OfertaIndex As Integer 'UserIndex del usuario con mayor oferta
    
    OfertaMayor As Long 'Oferta que vale la pena
    ValorBase As Long 'Valor base del Item
    
    Objeto As Obj 'Objeto
    
    Tiempo As Byte 'Tiempo de Subasta
End Type
    
Public Subasta As c_subasta

Public Sub Init_Subastas()
    With Subasta
        .Actual = False
        .UserIndex = 0
        .OfertaIndex = 0
        .ValorBase = 0
        .OfertaMayor = 0
        .Tiempo = 0
    End With
End Sub

Public Sub Iniciar_Subasta(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal ValorBase As Long)
        
    With Subasta
        If .Actual Then
            Call WriteConsoleMsg(UserIndex, "Ya hay una subasta, tenés que esperar a que termine para inciar una nueva subasta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            frmMain.Subasta.Enabled = True
            
            .Actual = True
            .UserIndex = UserIndex
            
            .OfertaIndex = 0
            .ValorBase = Val(ValorBase)
            .OfertaMayor = Val(.ValorBase)
            
            .Objeto.Amount = Amount
            .Objeto.index = ObjIndex
            
            .Tiempo = 3 '3 Minutos de Subasta
                        
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.UserIndex).Name & " está subastando " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & ".", FontTypeNames.FONTTYPE_INFO))
            
            Call EraseObj(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .Objeto.Amount)
            
            Exit Sub
        End If
    End With
End Sub

Public Sub Revisar_Subasta(ByVal UserIndex As Integer)
    With Subasta
        If UserIndex = .OfertaIndex Then
            .OfertaIndex = -1
        End If
        
        If UserIndex = .UserIndex > 0 Then
            .UserIndex = -1
        End If
    End With
End Sub

Public Sub Ofertar_Subasta(ByVal UserIndex As Integer, Oferta As Long)
    With Subasta
        If .Actual = False Then
            Call WriteConsoleMsg(UserIndex, "En este momento no hay ninguna subasta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        ElseIf .UserIndex = UserIndex Then
            Exit Sub
            
        ElseIf UserList(UserIndex).Stats.Gld < Oferta Then
            Call WriteConsoleMsg(UserIndex, "No tenés " & Oferta & " monedas de oro para ofertar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        ElseIf Oferta <= .OfertaMayor + .OfertaMayor * 0.05 Then
            Call WriteConsoleMsg(UserIndex, "Tu oferta debe superar por al menos un 5 por ciento la oferta de " & Val(.OfertaMayor) & " de " & UserList(.OfertaIndex).Name & ". En este momento la oferta mínima es de " & .OfertaMayor + .OfertaMayor * 0.1 & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        ElseIf .OfertaIndex = 0 And Oferta <= .ValorBase Then
            Call WriteConsoleMsg(UserIndex, "Tu oferta es menor al valor inicial de " & Val(.ValorBase) & ".", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        
        ElseIf .OfertaIndex > 0 Then
            UserList(.OfertaIndex).Stats.Gld = UserList(.OfertaIndex).Stats.Gld + Val(.OfertaMayor)
            Call WriteUpdateGold(.OfertaIndex)
        End If
            
        .OfertaIndex = UserIndex
        .OfertaMayor = Oferta
        
        UserList(.OfertaIndex).Stats.Gld = UserList(.OfertaIndex).Stats.Gld - Val(.OfertaMayor)
        Call WriteUpdateGold(.OfertaIndex)
        
        If .Tiempo = 1 Then
            .Tiempo = Tiempo + 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.OfertaIndex).Name & " ofertó " & .OfertaMayor & " monedas de oro. La subasta se alargó en 1 minuto.", FontTypeNames.FONTTYPE_INFO))
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.OfertaIndex).Name & " ofertó " & .OfertaMayor & " monedas de oro.", FontTypeNames.FONTTYPE_INFO))
        End If
    End With
End Sub

Public Sub Actualizar_Subasta()
    With Subasta
        If .Actual = False Then
            Exit Sub
        Else
            .Tiempo = .Tiempo - 1

            If .Tiempo < 1 Then
                Call Termina_Subasta
            Else
                If .UserIndex <> -1 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.UserIndex).Name & " está subastando " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & ". La mejor oferta es de " & .OfertaMayor & " monedas de oro. La subasta seguirá por " & .Tiempo & IIf(.Tiempo > 1, " minutos.", " minuto."), FontTypeNames.FONTTYPE_INFO))
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] Se está subastando " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguirá por " & .Tiempo & IIf(.Tiempo > 1, " minutos.", " minuto."), FontTypeNames.FONTTYPE_INFO))
                End If
            End If
        End If
    End With
End Sub

Public Sub Termina_Subasta()
    With Subasta
        If .OfertaIndex = 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] La subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & " terminó sin ninguna oferta.", FontTypeNames.FONTTYPE_INFO))
       
            If .UserIndex <> -1 Then
                Call MeterEnInventario(.UserIndex, .Objeto)
            End If
            
            .Actual = False
            .UserIndex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
        Else
            If .OfertaIndex <> -1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.OfertaIndex).Name & " ganó la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & " por " & .OfertaMayor & " monedas de oro.", FontTypeNames.FONTTYPE_INFO))
    
                If MeterEnInventario(.OfertaIndex, .Objeto) Then
                    Call WriteConsoleMsg(.OfertaIndex, "Felicitaciones, ganaste la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.index).Name & " por " & .OfertaMayor & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If .UserIndex <> -1 Then
                UserList(.UserIndex).Stats.Gld = UserList(.UserIndex).Stats.Gld + Val(.OfertaMayor)
                Call WriteUpdateGold(.UserIndex)
            End If
            
            .Actual = False
            .UserIndex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
            
            frmMain.Subasta.Enabled = False
        End If
    End With
End Sub
