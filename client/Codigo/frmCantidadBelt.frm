VERSION 5.00
Begin VB.Form frmCantidadBelt 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   1665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   0
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   240
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "1"
      Top             =   560
      Width           =   1185
   End
   Begin VB.Shape CantidadBorder 
      BorderColor     =   &H00C0FFFF&
      Height          =   345
      Left            =   345
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Todo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Todo"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Aceptar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tirar"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmCantidadBelt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub Aceptar_Click()

    If LenB(Cantidad.Text) > 0 Then
        Dim Cuanto As Long
        
        Cuanto = Val(Replace(Cantidad.Text, ".", vbNullString))
        
        If Cuanto > 0 And Cuanto < 100001 Then
            Call WriteDropGold(Cuanto)
            'Call Audio.Play(SND_DROP_GOLD)
        End If
    End If
    
    Unload Me
End Sub

Private Sub Aceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Aceptar.ForeColor = &H80FFFF
End Sub

Private Sub Cantidad_Change()

On Error GoTo ErrHandler

    If LenB(Cantidad.Text) < 1 Then
        Exit Sub
    End If
    
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For i = 1 To Len(Cantidad.Text)
        CharAscii = Asc(mid$(Cantidad.Text, i, 1))
        
        If (CharAscii > 47 And CharAscii < 58) Or CharAscii = 46 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    If tempstr <> Cantidad.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        Cantidad.Text = tempstr
    End If
    
    If Val(Replace(tempstr, ".", vbNullString)) > UserGld Then
        If UserGld < 100001 Then
            tempstr = CStr(UserGld)
        Else
            tempstr = "100000"
        End If
        
    ElseIf Val(Replace(tempstr, ".", vbNullString)) > 100000 Then
        tempstr = "100000"
    End If

    Cantidad.Text = PonerPuntos(Val(Replace(tempstr, ".", vbNullString)))
    Cantidad.SelStart = Len(Cantidad.Text)

    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Cantidad.Text = "1"
    Cantidad.SelStart = Len(Cantidad.Text)
End Sub

Private Sub Cantidad_KeyUp(KeyCode As Integer, Shift As Integer)
    Cantidad.SelStart = Len(Cantidad.Text)
End Sub

Private Sub Cantidad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        
        LockWindowUpdate Cantidad.hWnd
        
        Cantidad.Enabled = False
        
        Unload Me
    End If
End Sub

Private Sub Cantidad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Aceptar.ForeColor = &HC0FFFF
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        
    ElseIf KeyCode = vbKeyReturn Then
        If LenB(Cantidad.Text) > 0 Then
            Dim Cuanto As Long
            
            Cuanto = Val(Replace(Cantidad.Text, ".", vbNullString))
            
            If Cuanto > 0 And Cuanto < 100001 Then
                Call WriteDropGold(Cuanto)
                'Call Audio.Play(SND_DROP_GOLD)
            End If
        End If
        
        Unload Me
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Aceptar.ForeColor = &HC0FFFF
End Sub

Private Sub Form_Load()
    'Dim x As Long
    'Dim y As Long
    'Dim n As Long
    
    'x = Width / Screen.TwipsPerPixelX
    'y = Height / Screen.TwipsPerPixelY
    
    'set the corner angle by changing the value of 'n'
    'n = 25
    
    'SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True
    
    Call Make_Transparent_Form(hWnd, 210)
    
    Cantidad.Text = "1"
    
    Cantidad.SelStart = Len(Cantidad.Text)
    
    Picture = LoadPicture(GrhPath & "Cantidad.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    Else
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
    HideCaret Cantidad.hWnd
End Sub
