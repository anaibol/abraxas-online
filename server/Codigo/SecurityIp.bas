Attribute VB_Name = "SecurityIp"
Option Explicit

'General_IpSecurity.Bas - Maneja la seguridad de las IPs

Private IpTables()      As Long 'USAMOS 2 LONGS: UNO DE LA IP, SEGUIDO DE UNO DE LA INFO
Private EntrysCounter   As Long
Private MaxValue        As Long
Private Multiplicado    As Long 'Cuantas veces multiplike el EntrysCounter para que me entren?
Private Const IntervaloEntreConexiones As Long = 500

'
'Declaraciones para Maximas conexiones por usuario
'Agregado por EL OSO
Private MaxConTables()      As Long
Private MaxConTablesEntry   As Long     'puntero a la ultima insertada

Private Const LIMITECONEXIONESxIP As Long = 10

Private Enum e_SecurityIpTabla
    IP_INTERVALOS = 1
    IP_LIMITECONEXIONES = 2
End Enum

Public Sub InitIpTables(ByVal OptCountersValue As Long)
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: EL OSO 21\01\06. Soporte para MaxConTables
'
'*************************************************  *************
    EntrysCounter = OptCountersValue
    Multiplicado = 1

    ReDim IpTables(EntrysCounter * 2) As Long
    MaxValue = 0

    ReDim MaxConTables(Declaraciones.MaxPoblacion * 2 - 1) As Long
    MaxConTablesEntry = 0

End Sub

'
'
'FUNCIONES PARA INTERVALOS'
'
'

Public Sub IpSecurityMantenimientoLista()
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'
'
'*************************************************  *************
    'Las borro todas cada 1 hora, asi se "renuevan"
    EntrysCounter = EntrysCounter \ Multiplicado
    Multiplicado = 1
    ReDim IpTables(EntrysCounter * 2) As Long
    MaxValue = 0
End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal Ip As Long) As Boolean
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'
'
'*************************************************  *************
Dim IpTableIndex As Long
    

    IpTableIndex = FindTableIp(Ip, IP_INTERVALOS)
    
    If IpTableIndex > 1 Then
        If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= GetTickCount Then   'No está saturando de connects?
            IpTables(IpTableIndex + 1) = GetTickCount
            IpSecurityAceptarNuevaConexion = True
            Debug.Print "CONEXION ACEPTADA"
            Exit Function
        Else
            IpSecurityAceptarNuevaConexion = False

            Debug.Print "CONEXION NO ACEPTADA"
            Exit Function
        End If
    Else
        IpTableIndex = Not IpTableIndex
        AddNewIpIntervalo Ip, IpTableIndex
        IpTables(IpTableIndex + 1) = GetTickCount
        IpSecurityAceptarNuevaConexion = True
        Exit Function
    End If

End Function


Private Sub AddNewIpIntervalo(ByVal Ip As Long, ByVal Index As Long)
'*************************************************  *************
'Author: Lucio N. Tourrilhes (DuNga)
'
'
'*************************************************  *************
    '2) Pruebo si hay espacio, sino agrando la lista
    If MaxValue + 1 > EntrysCounter Then
        EntrysCounter = EntrysCounter \ Multiplicado
        Multiplicado = Multiplicado + 1
        EntrysCounter = EntrysCounter * Multiplicado
        
        ReDim Preserve IpTables(EntrysCounter * 2) As Long
    End If
    
    '4) Corro todo el array para arriba
    Call CopyMemory(IpTables(Index + 2), IpTables(Index), (MaxValue - Index * 0.5) * 8)   '*4 (peso del long) * 2(Cantidad de elementos por c\u)
    IpTables(Index) = Ip
    
    '3) Subo el indicador de el Maximo valor almacenado y listo :)
    MaxValue = MaxValue + 1
End Sub

'
'
'FUNCIONES PARA LIMITES X IP'
'
'

Public Function IPSecuritySuperaLimiteConexiones(ByVal Ip As Long) As Boolean
Dim IpTableIndex As Long

    IpTableIndex = FindTableIp(Ip, IP_LIMITECONEXIONES)
    
    If IpTableIndex > 1 Then
        
        If MaxConTables(IpTableIndex + 1) < LIMITECONEXIONESxIP Then
            LogIP ("Agregamos conexion a " & Ip & " iptableIndex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "suma conexion a " & Ip & " total " & MaxConTables(IpTableIndex + 1) + 1
            MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
            IPSecuritySuperaLimiteConexiones = False
        Else
            LogIP ("rechazamos conexion de " & Ip & " iptableIndex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "rechaza conexion a " & Ip
            IPSecuritySuperaLimiteConexiones = True
        End If
    Else
        IPSecuritySuperaLimiteConexiones = False
        If MaxConTablesEntry < Declaraciones.MaxPoblacion Then  'si hay espacio..
            IpTableIndex = Not IpTableIndex
            AddNewIpLimiteConexiones Ip, IpTableIndex    'iptableIndex es donde lo agrego
            MaxConTables(IpTableIndex + 1) = 1
        Else
            Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")
        End If
    End If

End Function

Private Sub AddNewIpLimiteConexiones(ByVal Ip As Long, ByVal Index As Long)
    'Debug.Print "agrega conexion a " & ip
    'Debug.Print "(Declaraciones.MaxPoblacion - Index) = " & (Declaraciones.MaxPoblacion - Index)
    '4) Corro todo el array para arriba
    'Call CopyMemory(MaxConTables(Index + 2), MaxConTables(Index), (MaxConTablesEntry - Index * 0.5) * 8)    '*4 (peso del long) * 2(Cantidad de elementos por c/u)
    'MaxConTables(Index) = ip

    '3) Subo el indicador de el Maximo valor almacenado y listo :)
    'MaxConTablesEntry = MaxConTablesEntry + 1

    Debug.Print "agrega conexion a " & Ip
    Debug.Print "(Declaraciones.MaxPoblacion - Index) = " & (Declaraciones.MaxPoblacion - Index)
    Debug.Print "Agrega conexion a nueva Ip " & Ip
    '4) Corro todo el array para arriba
    Dim temp() As Long
    ReDim temp((MaxConTablesEntry - Index * 0.5) * 2) As Long  'VB no deja inicializar con rangos variables...
    Call CopyMemory(temp(0), MaxConTables(Index), (MaxConTablesEntry - Index * 0.5) * 8)    '*4 (peso del long) * 2(Cantidad de elementos por c/u)
    Call CopyMemory(MaxConTables(Index + 2), temp(0), (MaxConTablesEntry - Index * 0.5) * 8)    '*4 (peso del long) * 2(Cantidad de elementos por c/u)
    MaxConTables(Index) = Ip

    '3) Subo el indicador de el Maximo valor almacenado y listo :)
    MaxConTablesEntry = MaxConTablesEntry + 1

End Sub

Public Sub IpRestarConexion(ByVal Ip As Long)
Dim key As Long
    Debug.Print "resta conexion a " & Ip
    
    key = FindTableIp(Ip, IP_LIMITECONEXIONES)
    
    If key > 1 Then
        If MaxConTables(key + 1) > 0 Then
            MaxConTables(key + 1) = MaxConTables(key + 1) - 1
        End If
        Call LogIP("restamos conexion a " & Ip & " key=" & key & ". Conexiones: " & MaxConTables(key + 1))
        If MaxConTables(key + 1) < 1 Then
            'la limpiamos
            Call CopyMemory(MaxConTables(key), MaxConTables(key + 2), (MaxConTablesEntry - (key * 0.5) + 1) * 8)
            MaxConTablesEntry = MaxConTablesEntry - 1
        End If
    Else 'Key < 1
        Call LogIP("restamos conexion a " & Ip & " key=" & key & ". NEGATIVO!!")
        'LogCriticEvent "SecurityIp.IpRestarconexion obtuvo un valor negativo en key"
    End If
End Sub

'FUNCIONES GENERALES'

Private Function FindTableIp(ByVal Ip As Long, ByVal Tabla As e_SecurityIpTabla) As Long
Dim First As Long
Dim Last As Long
Dim Middle As Long
    
    Select Case Tabla
        Case e_SecurityIpTabla.IP_INTERVALOS
            First = 0
            Last = MaxValue
            Do While First <= Last
                Middle = (First + Last) * 0.5
                
                If (IpTables(Middle * 2) < Ip) Then
                    First = Middle + 1
                ElseIf (IpTables(Middle * 2) > Ip) Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (Middle * 2)
        
        Case e_SecurityIpTabla.IP_LIMITECONEXIONES
            
            First = 0
            Last = MaxConTablesEntry

            Do While First <= Last
                Middle = (First + Last) * 0.5

                If MaxConTables(Middle * 2) < Ip Then
                    First = Middle + 1
                ElseIf MaxConTables(Middle * 2) > Ip Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function
                End If
            Loop
            FindTableIp = Not (Middle * 2)
    End Select
End Function



Public Function DumpTables()
Dim i As Integer

    For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
        Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
    Next i

End Function
