Attribute VB_Name = "mdlCOmercioConUsuario"
Option Explicit

Private Const Max_ORO_LOGUEABLE As Long = 50000
Private Const Max_OBJ_LOGUEABLE As Long = 1000

Public Const Max_OFFER_SLOTS As Byte = 30 '20
Public Const GOLD_OFFER_SLOT As Byte = Max_OFFER_SLOTS + 1

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To Max_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    
    Cant(1 To Max_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)

    On Error GoTo errhandler
    
    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = True
    
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = True
    
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
        
    End If
    
    Call FlushBuffer(Destino)
    
    Exit Sub
errhandler:
        Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

Public Sub EnviarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte)

    Dim ObjIndex As Integer
    Dim ObjAmount As Long
    
    With UserList(UserIndex)
        If OfferSlot = GOLD_OFFER_SLOT Then
            ObjIndex = iORO
            ObjAmount = UserList(.ComUsu.DestUsu).ComUsu.GoldAmount
        Else
            ObjIndex = UserList(.ComUsu.DestUsu).ComUsu.Objeto(OfferSlot)
            ObjAmount = UserList(.ComUsu.DestUsu).ComUsu.Cant(OfferSlot)
        End If
    End With
   
    Call WriteChangeUserTradeSlot(UserIndex, OfferSlot, ObjIndex, ObjAmount)
    Call FlushBuffer(UserIndex)

End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)

    Dim i As Long
    
    With UserList(UserIndex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.Confirmo = False
        .ComUsu.DestUsu = 0
        
        For i = 1 To Max_OFFER_SLOTS
            .ComUsu.Cant(i) = 0
            .ComUsu.Objeto(i) = 0
        Next i
        
        .ComUsu.GoldAmount = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)

    Dim TradingObj As Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean
    Dim OfferSlot As Byte

    UserList(UserIndex).ComUsu.Acepto = True
    
    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Exit Sub
    End If
    
    'Envio los Items a quien corresponde
    For OfferSlot = 1 To Max_OFFER_SLOTS + 1
        
        'Items del 1er usuario
        With UserList(UserIndex)
            'Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                'Quito la Cantidad de oro ofrecida
                .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                'Log
                If .ComUsu.GoldAmount > Max_ORO_LOGUEABLE Then
                    Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                End If
                
                Call WriteUpdateGold(UserIndex)
                'Se la doy al otro
                UserList(OtroUserIndex).Stats.Gld = UserList(OtroUserIndex).Stats.Gld + .ComUsu.GoldAmount
                'Update Otro Usuario
                Call WriteUpdateGold(OtroUserIndex)
                
            'Le pasa lo ofertado de los slots con Items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.Index = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.Cant(OfferSlot)
                
                'Quita el objeto y se lo da al otro
                If Not MeterEnInventario(OtroUserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj, , UserIndex)
                End If
            
                Call QuitarObjetos(TradingObj.Index, TradingObj.Amount, UserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07\09\07
                If ObjData(TradingObj.Index).Log = 1 Then
                    Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.Index).Name)
                End If
            
                'Es mucha cantidad?
                If TradingObj.Amount > Max_OBJ_LOGUEABLE Then
                'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.Index).NoLog <> 1 Then
                        Call LogDesarrollo(UserList(OtroUserIndex).Name & " le pasó en comercio seguro a " & .Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.Index).Name)
                    End If
                End If
            End If
        End With
        
        'Items del 2do usuario
        With UserList(OtroUserIndex)
            'Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                'Quito la Cantidad de oro ofrecida
                .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                'Log
                If .ComUsu.GoldAmount > Max_ORO_LOGUEABLE Then
                    Call LogDesarrollo(.Name & " soltó oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                End If
                
                'Update Usuario
                Call WriteUpdateGold(OtroUserIndex)
                'y se la doy al otro
                UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .ComUsu.GoldAmount
                
                If .ComUsu.GoldAmount > Max_ORO_LOGUEABLE Then
                    Call LogDesarrollo(UserList(UserIndex).Name & " recibió oro en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.GoldAmount)
                End If
                
                'Update Otro Usuario
                Call WriteUpdateGold(UserIndex)
                
            'Le pasa la oferta de los slots con Items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.Index = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.Cant(OfferSlot)
                
                'Quita el objeto y se lo da al otro
                If Not MeterEnInventario(UserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, TradingObj, , UserIndex)
                End If
            
                Call QuitarObjetos(TradingObj.Index, TradingObj.Amount, OtroUserIndex)
                
                'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07\09\07
                If ObjData(TradingObj.Index).Log = 1 Then
                    Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.Index).Name)
                End If
            
                'Es mucha cantidad?
                If TradingObj.Amount > Max_OBJ_LOGUEABLE Then
                'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.Index).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pasó en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.Index).Name)
                    End If
                End If
            End If
        End With
        
    Next OfferSlot

    'End Trade
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
 
End Sub

Public Sub AgregarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
'Adds gold or Items to the user's offer

    If PuedeSeguirComerciando(UserIndex) Then
        With UserList(UserIndex).ComUsu
            'Si ya confirmo su oferta, no puede cambiarla!
            If Not .Confirmo Then
                If IsGold Then
                'Agregamos (o quitamos) mas oro a la oferta
                    .GoldAmount = .GoldAmount + Amount
                    
                    'Imposible que pase, pero por las dudas..
                    If .GoldAmount < 0 Then
                        .GoldAmount = 0
                    End If
                Else
                'Agreamos (o quitamos) el Item y su Cantidad en el slot correspondiente
                    'Si es 0 estoy modificando la Cantidad, no agregando
                    If ObjIndex > 0 Then
                        .Objeto(OfferSlot) = ObjIndex
                    End If
                    
                    .Cant(OfferSlot) = .Cant(OfferSlot) + Amount
                    
                    'Quitó todos los Items de ese tipo
                    If .Cant(OfferSlot) < 1 Then
                        'Removemos el objeto para evitar conflictos
                        .Objeto(OfferSlot) = 0
                        .Cant(OfferSlot) = 0
                    End If
                End If
            End If
        End With
    End If

End Sub

Public Function PuedeSeguirComerciando(ByVal UserIndex As Integer) As Boolean
'Validates wether the conditions for the commerce to keep going are satisfied

Dim OtroUserIndex As Integer
Dim ComercioInvalido As Boolean

With UserList(UserIndex)
    'Usuario válido?
    If .ComUsu.DestUsu < 1 Or .ComUsu.DestUsu > MaxPoblacion Then
        ComercioInvalido = True
    End If
    
    OtroUserIndex = .ComUsu.DestUsu
    
    If Not ComercioInvalido Then
        'Estan logueados?
        If UserList(OtroUserIndex).flags.Logged Or .flags.Logged Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        'Se estan comerciando el uno al otro?
        If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        'El nombre del otro es el mismo que al que le comercio?
        If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        'Mi nombre  es el mismo que al que el le comercia?
        If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
            ComercioInvalido = True
        End If
    End If
    
    If Not ComercioInvalido Then
        'Esta vivo?
        If UserList(OtroUserIndex).Stats.Muerto Then
            ComercioInvalido = True
        End If
    End If
    
    'Fin del comercio
    If ComercioInvalido Then
        Call FinComerciarUsu(UserIndex)
        
        If OtroUserIndex < 1 Or OtroUserIndex > MaxPoblacion Then
            Call FinComerciarUsu(OtroUserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
        
        Exit Function
    End If
End With

PuedeSeguirComerciando = True

End Function


