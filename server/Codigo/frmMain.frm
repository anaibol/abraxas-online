VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abraxas"
   ClientHeight    =   1785
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Subasta 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   120
   End
   Begin VB.Timer TIMER_PET_AI 
      Enabled         =   0   'False
      Interval        =   220
      Left            =   2400
      Top             =   600
   End
   Begin VB.Timer SaveTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   120
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   480
      Top             =   60
   End
   Begin VB.CheckBox SUPERLOG 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "log"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   540
   End
   Begin VB.Timer NpcAtaca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1920
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1935
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "BroadCast"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label LblDataSent 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enviado (bytes/s):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3165
      TabIndex        =   11
      Top             =   270
      Width           =   1560
   End
   Begin VB.Label LblDataReceived 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recibido (bytes/s):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3165
      TabIndex        =   10
      Top             =   30
      Width           =   1620
   End
   Begin VB.Label Escuch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Poblacion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Población: -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Abraxas"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTIp As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_Message = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, TIp As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = Id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTIp = TIp & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Public Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxPoblacion
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.Logged Then
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.Logged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the Message right away
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call CerrarUsuario(iUserIndex)
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
    Static centinelSecs As Byte
    
    centinelSecs = centinelSecs + 1
    
    If centinelSecs = 5 Then
        'Every 5 seconds, we try to call the player's attention so it will report the code.
        Call modCentinela.CallUserAttention
        
        centinelSecs = 0
    End If
    
    frmMain.LblDataSent.Caption = "Enviado (bytes/s): " & CStr(DataSent)
    
    frmMain.LblDataReceived.Caption = "Recibido (bytes/s): " & CStr(DataReceived)
    
    DataSent = 0
    
    DataReceived = 0
    
    Call PasarSegundo 'sistema de desconexion de 10 segs
    
    Exit Sub
errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
Resume Next

End Sub

Private Sub CMDDUMP_Click()
    On Error Resume Next
    
    Dim i As Integer
    For i = 1 To MaxPoblacion
        Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " Logged: " & UserList(i).flags.Logged)
    Next i
    
    Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)
End Sub

Private Sub Command1_Click()
    Call SendData(SendTarget.ToAll, 0, Msg_ShowMessageBox(BroadMsg.Text))
End Sub

Public Sub InitMain(ByVal f As Byte)
    If f = 1 Then
        Call mnuSystray_Click
    Else
        frmMain.Show
    End If
End Sub

Private Sub Command2_Click()
    Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                
                If hHook Then
                    UnhookWindowsHookEx hHook: hHook = 0
                End If
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next
    
    'Borramos el icono del systray
    Dim i As Integer
    Dim nid As NOTIFYICONDATA
    
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_Message Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, vbNullString)
    
    i = Shell_NotifyIconA(NIM_DELETE, nid)
        

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
            
    Call QuitarIconoSystray
    
    Call LimpiaWsApi
    
    Dim LoopC As Integer
    
    For LoopC = 1 To MaxPoblacion
        If UserList(LoopC).ConnID <> -1 Then
            Call CloseSocket(LoopC)
        End If
    Next
    
    'Log
    Dim N As Integer
    N = FreeFile
    Open App.Path & "/logs/main.log" For Append Shared As #N
    Print #N, Date & " " & Time & " server cerrado."
    Close #N
    
    End
    
    Set SonidosMapas = Nothing
    
    DB_Conn.Close

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror
    
    Call SonidosMapas.ReproducirSonidosDeMapas
    
    Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To MaxPoblacion 'LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '¿User válido?
                
                If .ConnIDValida And .flags.Logged Then
                    
                    Call DoTileEvents(iUserIndex)
                    
                    If .flags.Paralizado > 0 Then
                        Call EfectoParalisisUser(iUserIndex)
                    End If
                    
                    If .flags.Ceguera > 0 Or .flags.Estupidez > 0 Then
                        Call EfectoCegueEstu(iUserIndex)
                    End If
                    
                    If Not .Stats.Muerto Then
                    
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then
                            Call EfectoLava(iUserIndex)
                        End If
                        
                        If .flags.Desnudo Then
                            If (.flags.Privilegios And PlayerType.User) > 0 Then
                                Call EfectoFrio(iUserIndex)
                            End If
                        End If
                        
                        If .flags.Meditando Then
                            If (.flags.Privilegios And PlayerType.User) > 0 Then
                                Call DoMeditar(iUserIndex)
                            End If
                        End If
                        
                        If .flags.Envenenado > 0 Then
                            If (.flags.Privilegios And PlayerType.User) > 0 Then
                                Call EfectoVeneno(iUserIndex)
                            End If
                        End If
                        
                        If .flags.AdminInvisible < 1 Then
                            If .flags.Invisible > 0 Then
                                Call EfectoInvisibilidad(iUserIndex)
                            End If
                            If .flags.Oculto > 0 Then
                                Call DoPermanecerOculto(iUserIndex)
                            End If
                        End If
                        
                        If .flags.Mimetizado Then
                            Call EfectoMimetismo(iUserIndex)
                        End If
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex)
                        
                        If Not .flags.Desnudo Then
                            If .Stats.MinHam > 0 And .Stats.MinSed > 0 Then
                                If Lloviendo Then
                                    If Not Intemperie(iUserIndex) Then
                                        If Not .flags.Descansando Then
                                        'No esta descansando
                                        
                                            If .flags.Envenenado < 1 Then
                                                Call Sanar(iUserIndex, SanaIntervaloSinDescansar)
                                            End If
                                            
                                            Call RecStamina(iUserIndex, StaminaIntervaloSinDescansar)

                                        Else
                                        'Esta descansando
                                        
                                            Call Sanar(iUserIndex, SanaIntervaloDescansar)

                                            Call RecStamina(iUserIndex, StaminaIntervaloDescansar)
                                            
                                            'termina de descansar automaticamente
                                            If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                                Call WriteRestOK(iUserIndex)
                                                Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                                .flags.Descansando = False
                                            End If
                                        End If
                                    End If
                                Else
                                    If Not .flags.Descansando Then
                                    'No esta descansando
                                    
                                        Call Sanar(iUserIndex, SanaIntervaloSinDescansar)
                                        
                                        Call RecStamina(iUserIndex, StaminaIntervaloSinDescansar)
                                        
                                    Else
                                    'Esta descansando
                                        Call Sanar(iUserIndex, SanaIntervaloDescansar)

                                        Call RecStamina(iUserIndex, StaminaIntervaloDescansar)
                                        
                                        'Termina de descansar automaticamente
                                        
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Terminaste de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansando = False
                                        End If
                                        
                                    End If
                                End If
                            End If
                        End If
                        
                        If .Pets.NroALaVez > 0 Then
                            Call TiempoInvocacion(iUserIndex)
                        End If
                        
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'Logged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuCerrar_Click()

If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()

    On Error Resume Next
    
    If FileExist(App.Path & "/logs/connect.log", vbNormal) Then
        Kill App.Path & "/logs/connect.log"
    End If
    
    If FileExist(App.Path & "/logs/haciendo.log", vbNormal) Then
        Kill App.Path & "/logs/haciendo.log"
    End If
    
    If FileExist(App.Path & "/logs/stats.log", vbNormal) Then
        Kill App.Path & "/logs/stats.log"
    End If
    
    If FileExist(App.Path & "/logs/asesinatos.log", vbNormal) Then
        Kill App.Path & "/logs/asesinatos.log"
    End If
    
    If FileExist(App.Path & "/logs/hackAttemps.log", vbNormal) Then
        Kill App.Path & "/logs/hackAttemps.log"
    End If
    
    If Not FileExist(App.Path & "/logs/nokillwsapi.txt") Then
        If FileExist(App.Path & "/logs\wsapi.log", vbNormal) Then
            Kill App.Path & "/logs\wsapi.log"
        End If
    End If
End Sub

Private Sub mnuServidor_Click()
    frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()
    
    Dim i As Integer
    Dim S As String
    Dim nid As NOTIFYICONDATA
    
    S = "Servidor Abraxas"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_Message Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
        
    If WindowState <> vbMinimized Then
        WindowState = vbMinimized
    End If
    
    Visible = False

End Sub

Private Sub NpcAtaca_Timer()

On Error Resume Next

    Dim Npc As Long
    
    For Npc = 1 To LastNpc
        NpcList(Npc).CanAttack = 1
    Next Npc

End Sub

Private Sub packetResend_Timer()
'Attempts to resend to the user all data that may be enqueued.

On Error GoTo errhandler:
    Dim i As Long
    
    For i = 1 To MaxPoblacion
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

    Exit Sub

errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub

Private Sub TIMER_AI_Timer()

    Dim NpcIndex As Long
    Dim X As Integer
    Dim Y As Integer
    Dim UseAI As Integer
    Dim mapa As Integer
    Dim e_p As Integer
    
    If Not haciendoBK And Not EnPausa Then

        For NpcIndex = 1 To LastNpc
        
            With NpcList(NpcIndex)

                If .flags.NpcActive Then
                    If .flags.Paralizado > 0 Or .flags.Inmovilizado > 0 Then

                        If .Contadores.Paralisis > 0 Then
                            .Contadores.Paralisis = .Contadores.Paralisis - 1
                        Else
                            .flags.Paralizado = 0
                            .flags.Inmovilizado = 0
                            Call SendData(SendTarget.ToNpcArea, NpcIndex, Msg_SetParalized(.Char.CharIndex, 0))
                        End If
                        
                    Else
                        mapa = .Pos.Map
                        
                        If mapa > 0 Then
                            If .Movement <> TipoAI.Estatico Then
                                If .flags.Paralizado < 1 Then
                                    If .MaestroUser < 1 Then
                                        Call NpcAI(NpcIndex)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
            End With
            
        Next NpcIndex
    End If
    
End Sub

Private Sub TIMER_PET_AI_Timer()
    Call NpcPetAi
End Sub

Private Sub SaveTimer_Timer()

On Error Resume Next

    Dim Horario As String
    
    Dim Horas As Byte
    Dim Minutos As Byte
    Dim Segundos As Byte
    
    Dim TempTiempo As EstadoTiempo
    
    Horario = Format(Now, "hhmmss")
    
    Horas = Val(Left$(Horario, 2))
    Minutos = Val(mid$(Horario, 3, 2))
    
    Segundos = Val(Right$(Horario, 2))

    If Segundos = 0 Then
        Call ModAreas.AreasOptimizacion
    
        'Actualizamos el centinela
        Call modCentinela.PasarMinutoCentinela
        
        Call PurgarPenas
        Call CheckIdleUser
        
        Select Case Minutos
            
            Case 30
                Call RegistrarEstadisticas
                
            Case 0
                Call RegistrarEstadisticas
                Call LimpiarMundo
                Call GuardarUsuarios
                Minutos = 0
            
            Case 59
                Call SendData(SendTarget.ToAll, 0, Msg_ConsoleMsg("El mundo será limpiado en un minuto.", FontTypeNames.FONTTYPE_VENENO))
    
        End Select
        
        If Tiempo = 0 Or Minutos = 0 Then
            If Horas < 8 And Horas > 5 Then
                TempTiempo = Amanecer
            ElseIf Horas < 12 And Horas > 5 Then
                TempTiempo = Mañana
            ElseIf Horas < 15 And Horas > 5 Then
                TempTiempo = Mediodía
            ElseIf Horas < 18 And Horas > 5 Then
                TempTiempo = Tarde
            ElseIf Horas < 21 And Horas > 5 Then
                TempTiempo = Anochecer
            Else
                TempTiempo = Noche
            End If
            
            If TempTiempo <> Tiempo Then
                Tiempo = TempTiempo
                Call SendData(SendTarget.ToAll, 0, Msg_Weather())
            End If
        End If
        
        'If Tiempo = 0 Or Horas = 0 Then
        '    If Weekday(Now(), vbMonday) > 5 Then
        '    End If
        'End If
        
    End If
    
End Sub

Private Sub Subasta_Timer()
    Call Actualizar_Subasta
End Sub

Private Sub tLluvia_Timer()
On Error GoTo errhandler

    Dim iCount As Long
    If Lloviendo Then
       For iCount = 1 To LastUser
            Call EfectoLluvia(iCount)
       Next iCount
    End If
    
    Exit Sub
    
errhandler:
    Call LogError("tLluvia " & Err.Number & ": " & Err.description)
End Sub

Private Sub tLluviaEvent_Timer()

    Exit Sub
    
    On Error GoTo ErrorHandler
    Static MinutosLloviendo As Long
    Static MinutosSinLluvia As Long
    
    If Not Lloviendo Then
        MinutosSinLluvia = MinutosSinLluvia + 1
        If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
                If RandomNumber(1, 100) <= 2 Then
                    Lloviendo = True
                    MinutosSinLluvia = 0
                    Call SendData(SendTarget.ToAll, 0, Msg_RainToggle())
                End If
        ElseIf MinutosSinLluvia >= 1440 Then
                    Lloviendo = True
                    MinutosSinLluvia = 0
                    Call SendData(SendTarget.ToAll, 0, Msg_RainToggle())
        End If
    Else
        MinutosLloviendo = MinutosLloviendo + 1
        If MinutosLloviendo >= 5 Then
                Lloviendo = False
                Call SendData(SendTarget.ToAll, 0, Msg_RainToggle())
                MinutosLloviendo = 0
        Else
                If RandomNumber(1, 100) <= 2 Then
                    Lloviendo = False
                    MinutosLloviendo = 0
                    Call SendData(SendTarget.ToAll, 0, Msg_RainToggle())
                End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Call LogError("Error tLluviaTimer")

End Sub
