VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4065
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
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   3105
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
      ForeColor       =   &H00FFFFFF&
      Height          =   3180
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   450
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   540
      Width           =   3165
   End
   Begin VB.Image imgDetalles 
      Height          =   300
      Left            =   2040
      Tag             =   "1"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image imgCerrar 
      Height          =   300
      Left            =   840
      Tag             =   "1"
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Picture = LoadPicture(GrhPath & "ListaGuildas.jpg")
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarListaClanes.jpg", _
                                    GrhPath & "BotónCerrarRolloverListaClanes.jpg", _
                                    GrhPath & "BotónCerrarClickListaClanes.jpg", Me)

    Call cBotonDetalles.Initialize(imgDetalles, GrhPath & "BotónDetallesListaClanes.jpg", _
                                    GrhPath & "BotónDetallesRolloverListaClanes.jpg", _
                                    GrhPath & "BotónDetallesClickListaClanes.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub guildslist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
End Sub

Private Sub imgDetalles_Click()
    frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(GuildsList.List(GuildsList.ListIndex))
End Sub

Private Sub txtBuscar_Change()
    Call FiltrarListaClanes(txtBuscar.Text)
End Sub

Private Sub txtBuscar_GotFocus()
    With txtBuscar
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    If UBound(GuildNames) > 0 Then
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
    End If

End Sub
