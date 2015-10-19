VERSION 5.00
Begin VB.Form frmCommet 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "FrizQuadrata BT"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
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
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   2880
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgEnviar 
      Height          =   480
      Left            =   1080
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Private cBotonEnviar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public Nombre As String
Public T As TIPO
Public Enum TIPO
    ALIANZA = 1
    PAZ = 2
    rechazóPJ = 3
End Enum

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Call LoadBackGround
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonEnviar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonEnviar.Initialize(imgEnviar, GrhPath & "BotónEnviarSolicitud.jpg", _
                                    GrhPath & "BotónEnviarRolloverSolicitud.jpg", _
                                    GrhPath & "BotónEnviarClickSolicitud.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarSolicitud.jpg", _
                                    GrhPath & "BotónCerrarRolloverSolicitud.jpg", _
                                    GrhPath & "BotónCerrarClickSolicitud.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub imgEnviar_Click()

    If LenB(Text1) < 1 Then
    If T = PAZ Or T = ALIANZA Then
        MsgBox "Debes Redactar un mensaje solicitando la paz o alianza al líder de " & Nombre
    Else
        MsgBox "Debes indicar el motivo por el cual rechazas la membresía de " & Nombre
    End If
        
    Exit Sub
End If

If T = PAZ Then
    Call WriteGuildOfferPeace(Nombre, Replace(Text1, vbCrLf, "º"))
        
ElseIf T = ALIANZA Then
    Call WriteGuildOfferAlliance(Nombre, Replace(Text1, vbCrLf, "º"))
        
ElseIf T = rechazóPJ Then
    Call WriteGuildRejectNewMember(Nombre, Replace(Replace(Text1.Text, ",", " "), vbCrLf, " "))
    'Sacamos el char de la lista de aspirantes
    Dim i As Long
        
    For i = 0 To frmGuildLeader.Solicitudes.ListCount - 1
        If frmGuildLeader.Solicitudes.List(i) = Nombre Then
            frmGuildLeader.Solicitudes.RemoveItem i
            Exit For
        End If
    Next i
    
    Hide
    Unload frmCharInfo
End If
    
Unload Me

End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then
        Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)
    End If
End Sub

Private Sub LoadBackGround()

    Select Case T
        Case TIPO.ALIANZA
            Picture = LoadPicture(GrhPath & "VentanaPropuestaAlianza.jpg")
            
        Case TIPO.PAZ
            Picture = LoadPicture(GrhPath & "VentanaPropuestaPaz.jpg")
            
        Case TIPO.rechazóPJ
            Picture = LoadPicture(GrhPath & "VentanaMotivorechazó.jpg")
            
    End Select
    
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub
