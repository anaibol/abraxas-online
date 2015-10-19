Attribute VB_Name = "wskapiAO"
Option Explicit

'Modulo para manejar Winsock

'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).

#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotsock se nos come todo el uso del CPU.

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

Public SockListen As Long

Public Sub IniciaWsApi(ByVal hwndParent As Long)

    Call LogApiSock("IniciaWsApi")
    Debug.Print "IniciaWsApi"
    
    #If WSAPI_CREAR_LABEL Then
    hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", &H40000000, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
    #Else
    hWndMsg = hwndParent
    #End If 'WSAPI_CREAR_LABEL
    
    OldWProc = SetWindowLong(hWndMsg, -4, AddressOf WndProc)
    ActualWProc = GetWindowLong(hWndMsg, -4)
    
    Dim Desc As String
    Call StartWinsock(Desc)

End Sub

Public Sub LimpiaWsApi()

    Call LogApiSock("LimpiaWsApi")
    
    If WSAStartedUp Then
        Call EndWinsock
    End If
    
    If OldWProc > 0 Then
        SetWindowLong hWndMsg, -4, OldWProc
        OldWProc = 0
    End If
    
    #If WSAPI_CREAR_LABEL Then
        If hWndMsg > 0 Then
            DestroyWindow hWndMsg
        End If
    #End If

End Sub

Public Function BuscaSlotsock(ByVal S As Long) As Long

On Error GoTo hayerror
    BuscaSlotsock = WSAPISock2Usr.Item(CStr(S))
    Exit Function
Exit Function
hayerror:
    BuscaSlotsock = -1

End Function

Public Sub AgregaSlotsock(ByVal Sock As Long, ByVal Slot As Long)
    Debug.Print "AgregaSockSlot"
    
    If WSAPISock2Usr.Count > MaxPoblacion Then
        Call CloseSocket(Slot)
        Exit Sub
    End If
    
    WSAPISock2Usr.Add CStr(Slot), CStr(Sock)
End Sub

Public Sub BorraSlotsock(ByVal Sock As Long)
    Dim Cant As Long
    
    Cant = WSAPISock2Usr.Count
    On Error Resume Next
    WSAPISock2Usr.Remove CStr(Sock)
    
    Debug.Print "BorraSockSlot " & Cant & " -> " & WSAPISock2Usr.Count
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next

    Dim ret As Long
    Dim Tmp() As Byte
    Dim S As Long
    Dim E As Long
    Dim N As Integer
    Dim UltError As Long
    
    Select Case msg
        Case 1025
            S = wParam
            E = WSAGetSelectEvent(lParam)
            
            Select Case E
                Case FD_ACCEPT
                    If S = SockListen Then
                        Call EventoSockAccept(S)
                    End If
                
                Case FD_READ
                    N = BuscaSlotsock(S)
                    If N < 0 And S <> SockListen Then
                        'Call apiclosesocket(s)
                        Call WSApiCloseSocket(S)
                        Exit Function
                    End If
                    
                    'create appropiate sized buffer
                    ReDim Preserve Tmp(8192 - 1) As Byte
                    
                    ret = recv(S, Tmp(0), 8192, 0)
                    'Comparo por = 0 ya que esto es cuando se cierra
                    If ret < 0 Then
                        UltError = Err.LastDllError
                        If UltError = WSAEMSGSIZE Then
                            Debug.Print "WSAEMSGSIZE"
                            ret = 8192
                        Else
                            Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
                            Call LogApiSock("Error en Recv: N=" & N & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                            
                            'no hay q llamar a CloseSocket() directamente,
                            'ya q pueden abusar de algun error para
                            'desconectarse sin los 10segs. CREEME.
                            Call CloseSocketSL(N)
                            Call CerrarUsuario(N)
                            Exit Function
                        End If
                    ElseIf ret = 0 Then
                        Call CloseSocketSL(N)
                        Call CerrarUsuario(N)
                    End If
                    
                    ReDim Preserve Tmp(ret - 1) As Byte
                    
                    Call EventoSockRead(N, Tmp)
                
                Case FD_CLOSE
                    N = BuscaSlotsock(S)
                    If S <> SockListen Then
                        Call apiclosesocket(S)
                    End If
                    
                    If N > 0 Then
                        Call BorraSlotsock(S)
                        UserList(N).ConnID = -1
                        UserList(N).ConnIDValida = False
                        Call EventoSockClose(N)
                    End If
            End Select
        
        Case Else
            WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
    End Select
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna > 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal Slot As Integer, ByRef str As String) As Long
    Dim ret As String
    Dim Retorno As Long
    Dim data() As Byte
    
    ReDim Preserve data(Len(str) - 1) As Byte

    data = StrConv(str, vbFromUnicode)
    
    Retorno = 0
    
    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
        ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
        If ret < 0 Then
            ret = Err.LastDllError
            If ret = WSAEWOULDBLOCK Then
                
                'WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(str)
            End If
        End If
    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1
        End If
    End If
    
    WsApiEnviar = Retorno
End Function

Public Sub LogApiSock(ByVal str As String)

On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile 'obtenemos un canal
Open App.Path & "/logs\wsapi.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
'==========================================================
'USO DE LA API DE WINSOCK
'========================
    
    Dim NewIndex As Integer
    Dim ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    
    Tam = sockaddr_size
    
    '=============================================
    'SockID es en este caso es el socket de escucha,
    'a diferencia de socketwrench que es el nuevo
    'socket de la nueva conn
    
'Modificado por Maraxus
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
    ret = accept(SockID, sa, Tam)

    If ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
    'If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
    '    Call WSApiCloseSocket(NuevoSock)
    '    EXIT SUB
    'End If

    'If Ret = INVALID_SOCKET Then
    'If Err.LastDllError = 11002 Then
    'We couldn't decide if to accept or reject the connection
    'Force reject so we can get it out of the queue
    'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
    'Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
    'Else
    'i = Err.LastDllError
    'Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
    'EXIT SUB
    'End If
    'End If

    NuevoSock = ret
    
    'Seteamos el tamaño del buffer de entrada
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, 8192, 4) > 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    'Seteamos el tamaño del buffer de salida
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, 8192, 4) > 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If

    'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
    '    Call send(ByVal NuevoSock, ByVal vbNullString, ByVal 0, ByVal 0)
    '    Call WSApiCloseSocket(NuevoSock)
    '    EXIT SUB
    'End If

    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser 'Nuevo indice
    
    If NewIndex <= MaxPoblacion Then
        
        'Make sure both outgoing and incoming data buffers are clean
        Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
        Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)

        UserList(NewIndex).Ip = GetAscIP(sa.sin_addr)
        
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).Ip Then
                'Call apiclosesocket(NuevoSock)
                Call WriteErrorMsg(NewIndex, "Fuiste desterrado del mundo de Abraxas.")
                Call FlushBuffer(NewIndex)
                'Call SecurityIp.IpRestarConexion(sa.sin_addr)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        If NewIndex > LastUser Then
            LastUser = NewIndex
        End If
            
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDValida = True
        
        Call AgregaSlotsock(NuevoSock, NewIndex)
    Else
        Dim str As String
        Dim data() As Byte
        
        str = Protocol.PrepareMessageErrorMsg("Abraxas está lleno. Probá ingresar más tarde.")
        
        ReDim Preserve data(Len(str) - 1) As Byte
        
        data = StrConv(str, vbFromUnicode)
        
        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
    End If
    
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)

With UserList(Slot)
    
    Call .incomingData.WriteBlock(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(Slot)
    Else
        Exit Sub
    End If
End With

End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)

    Dim CentinelaIndex As Byte
    CentinelaIndex = UserList(Slot).flags.CentinelaIndex
        
    If CentinelaIndex <> 0 Then
        Call modCentinela.CentinelaUserLogout(CentinelaIndex)
    End If
    
    If UserList(Slot).flags.Logged Then
        Call CloseSocketSL(Slot)
        Call CerrarUsuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
End Sub

Public Sub WSApiReiniciarSockets()
Dim i As Long
    'Cierra el socket de escucha
    If SockListen > 1 Then
        Call apiclosesocket(SockListen)
    End If
    
    'Cierra todas las conexiones
    For i = 1 To MaxPoblacion
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If
        
        'Call ResetUserSlot(i)
    Next i
    
    For i = 1 To MaxPoblacion
        Set UserList(i).incomingData = Nothing
        Set UserList(i).outgoingData = Nothing
    Next i
    
    ReDim UserList(1 To MaxPoblacion)
    For i = 1 To MaxPoblacion
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
        
        Set UserList(i).incomingData = New clsByteQueue
        Set UserList(i).outgoingData = New clsByteQueue
    Next i
    
    LastUser = 1

    Poblacion = 0
    frmMain.Poblacion.Caption = "Población: " & Poblacion
    Call Base.OnlinePlayers
    
    Call LimpiaWsApi
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
    Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
    Call ShutDown(Socket, SD_BOTH)
End Sub
