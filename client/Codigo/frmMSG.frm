VERSION 5.00
Begin VB.Form frmMSG 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   45
   ClientWidth     =   2445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "FrizQuadrata BT"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
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
      Height          =   2280
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   1845
   End
   Begin VB.Image imgCerrar 
      Height          =   420
      Left            =   375
      Tag             =   "1"
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub Form_Deactivate()
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Visible = False
    List1.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Visible = False
    List1.Clear
End Sub

Private Sub Form_Load()
    List1.Clear
    
    'Picture = LoadPicture(GrhPath & "VentanaShowSos.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarShowSos.jpg", _
                                    GrhPath & "BotónCerrarRolloverShowSos.jpg", _
                                    GrhPath & "BotónCerrarClickShowSos.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Visible = False
    List1.Clear
End Sub

Private Sub list1_Click()
    Dim ind As Integer
    ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("-")))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu menU_usuario
    End If

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub mnuBorrar_Click()
    If List1.ListIndex < 0 Then
        Exit Sub
    End If
    
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSOSRemove(aux)
    'Call WriteSOSRemove(List1.List(List1.listIndex))
    
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub mnuIR_Click()
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteGoToChar(aux)
    'Call WriteGoToChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
    
End Sub

Private Sub mnutraer_Click()
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.ListIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.ListIndex), Asc("-"))))
    Call WriteSummonChar(aux)
    'Call WriteSummonChar(ReadField(1, List1.List(List1.listIndex), Asc("-")))
End Sub
