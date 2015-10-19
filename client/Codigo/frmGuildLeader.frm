VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFiltrarMiembros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Left            =   3075
      TabIndex        =   6
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtFiltrarClanes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtGuildNews 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   690
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3435
      Width           =   5475
   End
   Begin VB.ListBox Solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   705
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   195
      List            =   "frmGuildLeader.frx":0007
      TabIndex        =   2
      Top             =   5100
      Width           =   2595
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1380
      ItemData        =   "frmGuildLeader.frx":0018
      Left            =   3060
      List            =   "frmGuildLeader.frx":001A
      TabIndex        =   1
      Top             =   540
      Width           =   2595
   End
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1380
      ItemData        =   "frmGuildLeader.frx":001C
      Left            =   180
      List            =   "frmGuildLeader.frx":0023
      TabIndex        =   0
      Top             =   540
      Width           =   2595
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   6705
      Width           =   2775
   End
   Begin VB.Image imgPropuestasAlianzas 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   6195
      Width           =   2775
   End
   Begin VB.Image imgPropuestasPaz 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   5685
      Width           =   2775
   End
   Begin VB.Image imgEditarDesc 
      Height          =   495
      Left            =   3000
      Tag             =   "1"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Image imgActualizar 
      Height          =   390
      Left            =   150
      Tag             =   "1"
      Top             =   4230
      Width           =   5550
   End
   Begin VB.Image imgDetallesSolicitudes 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   6045
      Width           =   2655
   End
   Begin VB.Image imgDetallesMiembros 
      Height          =   375
      Left            =   3060
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgDetallesGuilda 
      Height          =   375
      Left            =   165
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgElecciones 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   1815
      TabIndex        =   3
      Top             =   6510
      Width           =   255
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_NEWS_LENGTH As Integer = 512
Private clsFormulario As clsFormMovementManager

Private cBotonElecciones As clsGraphicalButton
Private cBotonActualizar As clsGraphicalButton
Private cBotonDetallesGuilda As clsGraphicalButton
Private cBotonDetallesMiembros As clsGraphicalButton
Private cBotonDetallesSolicitudes As clsGraphicalButton
Private cBotonEditarDesc As clsGraphicalButton
Private cBotonPropuestasPaz As clsGraphicalButton
Private cBotonPropuestasAlianzas As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Picture = LoadPicture(GrhPath & "VentanaAdministrarGuilda.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonElecciones = New clsGraphicalButton
    Set cBotonActualizar = New clsGraphicalButton
    Set cBotonDetallesGuilda = New clsGraphicalButton
    Set cBotonDetallesMiembros = New clsGraphicalButton
    Set cBotonDetallesSolicitudes = New clsGraphicalButton
    Set cBotonEditarDesc = New clsGraphicalButton
    Set cBotonPropuestasPaz = New clsGraphicalButton
    Set cBotonPropuestasAlianzas = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonElecciones.Initialize(imgElecciones, GrhPath & "BotónElecciones.jpg", _
                                    GrhPath & "BotónEleccionesRollover.jpg", _
                                    GrhPath & "BotónEleccionesClick.jpg", Me)

    Call cBotonActualizar.Initialize(imgActualizar, GrhPath & "BotónActualizar.jpg", _
                                    GrhPath & "BotónActualizarRollover.jpg", _
                                    GrhPath & "BotónActualizarClick.jpg", Me)

    Call cBotonDetallesGuilda.Initialize(imgDetallesGuilda, GrhPath & "BotónDetallesAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesRolloverAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesClickAdministrarGuilda.jpg", Me)

    Call cBotonDetallesMiembros.Initialize(imgDetallesMiembros, GrhPath & "BotónDetallesAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesRolloverAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesClickAdministrarGuilda.jpg", Me)
                                    
    Call cBotonDetallesSolicitudes.Initialize(imgDetallesSolicitudes, GrhPath & "BotónDetallesAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesRolloverAdministrarGuilda.jpg", _
                                    GrhPath & "BotónDetallesClickAdministrarGuilda.jpg", Me)

    'Call cBotonEditarDesc.Initialize(imgEditarDesc, GrhPath & "BotónEditarDesc.jpg", _
                                    GrhPath & "BotónEditarDescRollover.jpg", _
                                    GrhPath & "BotónEditarDescClick.jpg", Me)


    Call cBotonPropuestasPaz.Initialize(imgPropuestasPaz, GrhPath & "BotónPropuestaPaz.jpg", _
                                    GrhPath & "BotónPropuestaPazRollover.jpg", _
                                    GrhPath & "BotónPropuestaPazClick.jpg", Me)

    Call cBotonPropuestasAlianzas.Initialize(imgPropuestasAlianzas, GrhPath & "BotónPropuestasAlianzas.jpg", _
                                    GrhPath & "BotónPropuestasAlianzasRollover.jpg", _
                                    GrhPath & "BotónPropuestasAlianzasClick.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarAdministrarGuilda.jpg", _
                                    GrhPath & "BotónCerrarRolloverAdministrarGuilda.jpg", _
                                    GrhPath & "BotónCerrarClickAdministrarGuilda.jpg", Me)


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub guildslist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgActualizar_Click()
    Dim k As String

    k = Replace(txtGuildNews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub imgDetallesGuilda_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(GuildsList.List(GuildsList.ListIndex))
End Sub

Private Sub imgDetallesMiembros_Click()
    If members.ListIndex = -1 Then
        Exit Sub
    End If
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))
End Sub

Private Sub imgDetallesSolicitudes_Click()
    If Solicitudes.ListIndex = -1 Then
        Exit Sub
    End If
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(Solicitudes.List(Solicitudes.ListIndex))
End Sub

Private Sub imgElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me
End Sub

Private Sub imgPropuestasAlianzas_Click()
    Call WriteGuildAlliancePropList
End Sub

Private Sub imgPropuestasPaz_Click()
    Call WriteGuildPeacePropList
End Sub

Private Sub members_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub solicitudes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub txtguildnews_Change()
    If Len(txtGuildNews.Text) > MAX_NEWS_LENGTH Then
        txtGuildNews.Text = Left$(txtGuildNews.Text, MAX_NEWS_LENGTH)
    End If
End Sub

Private Sub txtguildnews_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub txtFiltrarClanes_Change()
    Call FiltrarListaClanes(txtFiltrarClanes.Text)
End Sub

Private Sub txtFiltrarClanes_GotFocus()
    With txtFiltrarClanes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    With GuildsList
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        'Recorro los arrays
        For lIndex = 0 To UBound(GuildNames)
            'Si coincide con los patrones
            If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                'Lo agrego a la lista
                .AddItem GuildNames(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With

End Sub

Private Sub txtFiltrarMiembros_Change()
    Call FiltrarListaMiembros(txtFiltrarMiembros.Text)
End Sub

Private Sub txtFiltrarMiembros_GotFocus()
    With txtFiltrarMiembros
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)

    Dim lIndex As Long
    
    With members
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        'Recorro los arrays
        For lIndex = 0 To UBound(GuildMembers)
            'Si coincide con los patrones
            If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
                'Lo agrego a la lista
                .AddItem GuildMembers(lIndex)
            End If
        Next lIndex
        
        .Visible = True
    End With
End Sub


