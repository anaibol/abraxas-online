VERSION 5.00
Begin VB.Form frmPeaceProp 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
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
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   240
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4620
   End
   Begin VB.Image imgRechazar 
      Height          =   480
      Left            =   3840
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   2640
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgDetalle 
      Height          =   480
      Left            =   1440
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   240
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton

Public LastPressed As clsGraphicalButton


Private TipoProp As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadBackGround
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotónAceptarOferta.jpg", _
                                    GrhPath & "BotónAceptarRolloverOferta.jpg", _
                                    GrhPath & "BotónAceptarClickOferta.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarOferta.jpg", _
                                    GrhPath & "BotónCerrarRolloverOferta.jpg", _
                                    GrhPath & "BotónCerrarClickOferta.jpg", Me)

    Call cBotonDetalles.Initialize(imgDetalle, GrhPath & "BotónDetallesOferta.jpg", _
                                    GrhPath & "BotónDetallesRolloverOferta.jpg", _
                                    GrhPath & "BotónDetallesClickOferta.jpg", Me)

    Call cBotonRechazar.Initialize(imgRechazar, GrhPath & "BotónRechazarOferta.jpg", _
                                    GrhPath & "BotónRechazarRolloverOferta.jpg", _
                                    GrhPath & "BotónRechazarClickOferta.jpg", Me)


End Sub

Private Sub LoadBackGround()
    If TipoProp = TIPO_PROPUESTA.ALIANZA Then
        Picture = LoadPicture(GrhPath & "VentanaOfertaAlianza.jpg")
    Else
        Picture = LoadPicture(GrhPath & "VentanaOfertaPaz.jpg")
    End If
End Sub

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
    TipoProp = nValue
End Property

Private Sub imgAceptar_Click()

    If TipoProp = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
    End If
    
    Hide
    
    Unload Me

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDetalle_Click()
    If TipoProp = PAZ Then
        Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))
    End If
End Sub

Private Sub imgRechazar_Click()

    If TipoProp = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
    End If
    
    Hide
    
    Unload Me
End Sub
