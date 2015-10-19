VERSION 5.00
Begin VB.Form frmServidor 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Servidor"
   ClientHeight    =   4980
   ClientLeft      =   3450
   ClientTop       =   2115
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRecargarAdministradores 
      BackColor       =   &H0080C0FF&
      Caption         =   "Administradores"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer BackUp"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2775
      TabIndex        =   23
      Top             =   1605
      Width           =   1545
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guardar PJ's"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   22
      Top             =   1110
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2730
      TabIndex        =   21
      Top             =   2145
      Width           =   1590
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Boton Mágico"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2715
      TabIndex        =   20
      Top             =   2595
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Obj.dat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2805
      TabIndex        =   18
      Top             =   3045
      Width           =   1320
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Configurar intervalos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4650
      TabIndex        =   17
      Top             =   1140
      Width           =   2250
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Hechizos.dat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   16
      Top             =   1785
      Width           =   1560
   End
   Begin VB.CommandButton Command9 
      Caption         =   "NombresInvalidos.txt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   165
      TabIndex        =   15
      Top             =   3075
      Width           =   2190
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Trafico"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4635
      TabIndex        =   14
      Top             =   1425
      Width           =   2250
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Stats de los slots"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4635
      TabIndex        =   13
      Top             =   1665
      Width           =   2250
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Debug Npcs"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4635
      TabIndex        =   12
      Top             =   1905
      Width           =   2250
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Unban All (PELIGRO!)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4635
      TabIndex        =   11
      Top             =   2145
      Width           =   2250
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Unban All IPs (PELIGRO!)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4635
      TabIndex        =   10
      Top             =   2385
      Width           =   2250
   End
   Begin VB.CommandButton Command14 
      Caption         =   "MOTD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1515
      TabIndex        =   9
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Balance.dat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   8
      Top             =   2145
      Width           =   1305
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Servidor.ini"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   225
      TabIndex        =   7
      Top             =   2640
      Width           =   1260
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Npcs.dat"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1215
      TabIndex        =   6
      Top             =   1305
      Width           =   960
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Pausar el servidor"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4650
      TabIndex        =   5
      Top             =   2910
      Width           =   2250
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Administración"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4650
      TabIndex        =   4
      Top             =   3150
      Width           =   2250
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Debug UserList"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4650
      TabIndex        =   3
      Top             =   3390
      Width           =   2250
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Reset Listen"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2505
      TabIndex        =   2
      Top             =   4245
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5805
      TabIndex        =   0
      Top             =   4395
      Width           =   1230
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reset sockets"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   795
      TabIndex        =   1
      Top             =   4245
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   5535
      TabIndex        =   27
      Top             =   645
      Width           =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "BackUp"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   3090
      TabIndex        =   26
      Top             =   690
      Width           =   1200
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   525
      TabIndex        =   25
      Top             =   645
      Width           =   1470
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Panel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   420
      Left            =   3075
      TabIndex        =   24
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecargarAdministradores_Click()
    Call LoadAdministrativeUsers
End Sub

Private Sub Command1_Click()
    Call LoadOBJData
End Sub

Private Sub Command10_Click()
    frmTrafic.Show
End Sub

Private Sub Command11_Click()
    frmConID.Show
End Sub

Private Sub Command12_Click()
    frmDebugNpc.Show
End Sub

Private Sub Command14_Click()
    Call LoadMotd
End Sub

Private Sub Command15_Click()
On Error Resume Next

    Dim Fn As String
    Dim cad$
    Dim N As Integer, k As Integer
    
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe vbnullstringestoy DE acuerdovbnullstring entre comillas y con distición de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
    
        Fn = App.Path & "/logs/genteBanned.log"
        
        If FileExist(Fn, vbNormal) Then
            N = FreeFile
            Open Fn For Input Shared As #N
            Do While Not EOF(N)
                k = k + 1
                Input #N, cad$
                Call UnBan(cad$)
                
            Loop
            Close #N
            MsgBox "Se han habilitado " & k & " personajes."
            Kill Fn
        End If
    End If

End Sub

Private Sub Command16_Click()
    Call LoadSini
End Sub

Private Sub Command17_Click()
    Call CargaNpcsDat
End Sub

Private Sub Command18_Click()
    Me.MousePointer = 11
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Me.MousePointer = vbDefault
    MsgBox "Grabado de personajes OK!"
End Sub

Private Sub Command19_Click()
    Dim i As Long, N As Long
    
    Dim sENtrada As String
    
    sENtrada = InputBox("Escribe vbnullstringestoy DE acuerdovbnullstring sin comillas y con distición de mayusculas minusculas para desbanear a todos los personajes", "UnBan", "hola")
    If sENtrada = "estoy DE acuerdo" Then
        
        N = BanIps.Count
        For i = 1 To BanIps.Count
            BanIps.Remove 1
        Next i
        
        MsgBox "Se han habilitado " & N & " ipes"
    End If

End Sub

Private Sub Command2_Click()
frmServidor.Visible = False
End Sub

Private Sub Command20_Click()
    Call WSApiReiniciarSockets
End Sub

Private Sub Command21_Click()

    If EnPausa = False Then
        EnPausa = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Command21.Caption = "Reanudar el servidor"
    Else
        EnPausa = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Command21.Caption = "Pausar el servidor"
    End If

End Sub

Private Sub Command22_Click()
    Me.Visible = False
    frmAdmin.Show
End Sub

Private Sub Command23_Click()
    If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, "Apagar Magicamente") = vbYes Then
        Me.MousePointer = 11
        
        FrmStat.Show
       
        'WorldSave
        Call ES.DoBackUp
    
        'commit experiencia
        Call mdParty.ActualizaExperiencias
    
        'Guardar Pjs
        Call GuardarUsuarios
        
        'Chauuu
        Unload frmMain
    End If
End Sub

Private Sub Command26_Click()
    'Cierra el socket de escucha
    If SockListen > 1 Then
        Call apiclosesocket(SockListen)
    End If
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
End Sub

Private Sub Command27_Click()
    frmUserList.Show
End Sub

Private Sub Command28_Click()
    Call LoadBalance
End Sub

Private Sub Command3_Click()
    If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la pérdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
        Me.Visible = False
        Call General.Restart
    End If
End Sub

Private Sub Command4_Click()
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Show
    Call ES.DoBackUp
    Me.MousePointer = vbDefault
    MsgBox "WORLDSAVE OK!!"
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Reiniciando."
    End If
    
    FrmStat.Show
    
    If FileExist(App.Path & "/logs\errores.log", vbNormal) Then
        Kill App.Path & "/logs\errores.log"
    End If
    
    If FileExist(App.Path & "/logs/connect.log", vbNormal) Then
        Kill App.Path & "/logs/connect.log"
    End If
    
    If FileExist(App.Path & "/logs/hackAttemps.log", vbNormal) Then
        Kill App.Path & "/logs/hackAttemps.log"
    End If
    
    If FileExist(App.Path & "/logs/asesinatos.log", vbNormal) Then
        Kill App.Path & "/logs/asesinatos.log"
    End If
    
    If FileExist(App.Path & "/logs/resurrecciones.log", vbNormal) Then
        Kill App.Path & "/logs/resurrecciones.log"
    End If
    
    If FileExist(App.Path & "/logs/teleports.Log", vbNormal) Then
        Kill App.Path & "/logs/teleports.Log"
    End If
    
    Call apiclosesocket(SockListen)
    
    Dim LoopC As Integer
    
    For LoopC = 1 To MaxPoblacion
        Call CloseSocket(LoopC)
    Next
    
    LastUser = 0
    
    Poblacion = 0
    frmMain.Poblacion.Caption = "Población: " & Poblacion
    Call Base.OnlinePlayers
    
    Call FreeNpcs
    Call FreeCharIndexes
    
    Call LoadSini
    Call LoadOBJData
    
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
    
    If frmMain.Visible Then
        frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    End If

End Sub

Private Sub Command7_Click()
    FrmInterv.Show
End Sub

Private Sub Command8_Click()
    Call CargarHechizos
End Sub

Private Sub Command9_Click()
    Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
    frmServidor.Visible = False
End Sub

Private Sub Form_Load()
    Command20.Visible = True
    Command26.Visible = True
End Sub
