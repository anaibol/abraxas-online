VERSION 5.00
Begin VB.Form frmBanco 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3645
   ClientLeft      =   2130
   ClientTop       =   2685
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
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
      Height          =   225
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   16
      Text            =   "1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2865
      Left            =   3240
      Picture         =   "frmBancoObj.frx":0000
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Tag             =   "5"
      Top             =   390
      Width           =   2400
   End
   Begin VB.TextBox Cantidad2 
      Alignment       =   2  'Center
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
      Height          =   225
      Left            =   1020
      MaxLength       =   7
      TabIndex        =   15
      Text            =   "1"
      Top             =   2400
      Width           =   840
   End
   Begin VB.Image Transferir 
      Height          =   255
      Left            =   975
      Top             =   2700
      Width           =   930
   End
   Begin VB.Image Retirar 
      Height          =   345
      Left            =   240
      Top             =   2325
      Width           =   690
   End
   Begin VB.Image Depositar 
      Height          =   345
      Left            =   2040
      Top             =   2325
      Width           =   930
   End
   Begin VB.Shape CantidadBorder 
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   6  'Inside Solid
      Height          =   285
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Shape CantidadBorder2 
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   6  'Inside Solid
      Height          =   285
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblItemName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Item"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   270
      Left            =   360
      TabIndex        =   10
      Top             =   345
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblPuedeUsar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No lo podés usar."
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblValor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblUserBankGold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   1
      Left            =   2400
      Top             =   2640
      Width           =   570
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   0
      Left            =   2400
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label lblItemPrice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1455314"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   270
      Left            =   945
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblGuion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblDefense 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa:"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   360
      TabIndex        =   11
      Top             =   870
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblGuion2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   270
      Left            =   1455
      TabIndex        =   2
      Top             =   870
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblMinHit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   225
      Left            =   960
      TabIndex        =   7
      Top             =   870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMaxDef 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Top             =   870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMaxHit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDamage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Daño:"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   870
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblBoveda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bóveda"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   690
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblMinDef 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   225
      Left            =   1200
      TabIndex        =   13
      Top             =   870
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        
        If CharAscii > 47 And CharAscii < 58 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    If Val(tempstr) > MaxInvObjs Then
       tempstr = MaxInvObjs
    ElseIf Val(tempstr) < 1 Then
        tempstr = "1"
    End If
    
    If Val(tempstr) > MaxInvObjs Then
        tempstr = MaxInvObjs
    ElseIf Val(tempstr) < 1 Then
        tempstr = "1"
    End If
    
    tempstr = PonerPuntos(Val(tempstr))
    
    If tempstr <> Cantidad.Text Then
        Cantidad.Text = tempstr
        Cantidad.SelStart = Len(Cantidad.Text)
    End If
    
    If NpcInvSelSlot > 0 Then
        'El precio, cuando nos venden algo, lo tenemos que Redondear para arriba.
        lblItemPrice.Caption = PonerPuntos(CalculateBuyPrice(NpcInv(NpcInvSelSlot).Valor, Val(Cantidad.Text)))
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Cantidad.Text = "1"
    Cantidad.SelStart = Len(Cantidad.Text)
End Sub

Private Sub Cantidad_GotFocus()
    CantidadBorder.BorderColor = &H80FFFF
End Sub

Private Sub Cantidad_LostFocus()
    CantidadBorder.BorderColor = &H80FFFF
End Sub

Private Sub Cantidad2_GotFocus()
    CantidadBorder2.BorderColor = &H80FFFF
End Sub

Private Sub Cantidad2_LostFocus()
    CantidadBorder2.BorderColor = &H80FFFF
End Sub

Private Sub Cantidad2_Change()
On Error GoTo ErrHandler

    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim i As Long
    Dim tempstr As String
    Dim CharAscii As Integer
    
    For i = 1 To Len(Cantidad2.Text)
        CharAscii = Asc(mid$(Cantidad2.Text, i, 1))
        
        If CharAscii >= 48 And CharAscii <= 57 Then
            tempstr = tempstr & Chr$(CharAscii)
        End If
    Next i
    
    If tempstr <> Cantidad2.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        Cantidad2.Text = tempstr
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Cantidad2.Text = "1"
End Sub

Private Sub CmdMoverBov_Click(Index As Integer)

    If NpcInvSelSlot = 0 Then
        Exit Sub
    End If
    
    Select Case Index
        Case 1 'subir
            If NpcInvSelSlot <= 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("No podés mover el objeto en esa dirección.", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
        Case 0 'bajar
            If NpcInvSelSlot >= MaxBankSlots Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("No podés mover el objeto en esa dirección.", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
        End Select
    Call WriteMoveBankSlot(Index, NpcInvSelSlot)
End Sub

Private Sub Depositar_Click()

    If Not MainTimer.Check(TimersIndex.BuySell) Then
        Exit Sub
    End If
    
    If Val(Cantidad2.Text) <= 0 Or Val(Cantidad2.Text) > UserGld Or Not IsNumeric(Cantidad2.Text) Then
        Cantidad2.ForeColor = vbRed
        Exit Sub
    End If

    Cantidad2.ForeColor = &HC0FFFF

    Call WriteBankDepositGold(Val(Cantidad2.Text))
    
    UserBankGold = UserBankGold + Val(Cantidad2.Text)
    lblUserBankGold.Caption = PonerPuntos(UserBankGold)
    
    UserGld = UserGld - Val(Cantidad2.Text)
    frmMain.GldLbl.Caption = PonerPuntos(UserGld)
    
    Call InitGld(-Val(Cantidad2.Text))
End Sub

Private Sub Form_Load()
    'Dim x As Long
    'Dim y As Long
    'Dim n As Long
    
    'x = Width / Screen.TwipsPerPixelX
    'y = Height / Screen.TwipsPerPixelY
    
    'n = 25
    
    'SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, x, y, n, n), True
    
    Call Make_Transparent_Form(hWnd, 225)
    
    NpcInvSelSlot = 0
    
    Picture = LoadPicture(GrhPath & "Bóveda.jpg")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call Auto_Drag(hWnd)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If DragType = None Then
        lblItemName.Visible = False
        lblValor.Visible = False
        lblItemPrice.Visible = False
        lblPuedeUsar.Visible = False
        
        lblDamage.Visible = False
        lblMinHit.Visible = False
        lblMaxHit.Visible = False
                                
        lblDefense.Visible = False
        lblMinDef.Visible = False
        lblMaxDef.Visible = False
        
        lblGuion.Visible = False
        lblGuion2.Visible = False
        
        lblBoveda.Visible = True
    'End If
    
    If NpcTempSlot > 0 Then
        Dim Prev As Byte
        
        Prev = NpcTempSlot
        
        NpcTempSlot = 0
        
        If InvNpc(Prev).ObjIndex > 0 Then
            Call InvNpc.DrawSlot(Prev)
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Unload Me
        Comerciando = False
        Call WriteBankEnd
    End If
End Sub
Private Sub PicBancoInv_DblClick()

    If Not MainTimer.Check(TimersIndex.BuySell) Then
        Exit Sub
    End If
    
    If LenB(Cantidad.Text) < 1 Then
        Exit Sub
    End If
    
    Dim Cuanto As Integer
    
    Cuanto = Val(Replace(Cantidad.Text, ".", vbNullString))
                
    If NpcInvSelSlot > 0 Then
        If Cuanto > 0 And Cuanto <= MaxInvObjs Then
            If Cuanto > Banco(NpcInvSelSlot).Amount Then
                Cuanto = Banco(NpcInvSelSlot).Amount
            End If
            
            MousePointer = vbDefault
            Call WriteBankExtractItem(NpcInvSelSlot, Cuanto)
            Call Audio.Play(SND_CLICK)
        End If
    End If
End Sub
 
Private Sub PicBancoInv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Comerciando = False
        Call WriteBankEnd
    End If
End Sub

Private Sub PicBancoInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        Unload Me
        Comerciando = False
        Call WriteBankEnd
    End If
    
    'If Button = vbRightButton Then
    'If NpcInvSelSlot > 0 Then
    'If IsNumeric(Cantidad.Text) Then
    'If Cantidad.Text > 0 Then
    'MOUSEPOINTER = vbDefault

    'Call WriteBankExtractItem(NpcInvSelSlot, Cantidad.Text)
    'call Audio.Play(SND_CLICK)
    'End If
    'End If
    'End If
    'EXIT SUB
    'End If
    
    'If DragType = InventarioNpc Then
    'If NpcInvSelSlot > 0 Then
    'If IsNumeric(Cantidad.Text) Then
    'If Cantidad.Text > 0 Then
    'MOUSEPOINTER = vbDefault
    'Call WriteBankExtractItem(NpcInvSelSlot, Cantidad.Text)
    'call Audio.Play(SND_CLICK)
    'End If
    'End If
    'End If
    'DragType = None
    'MOUSEPOINTER = vbDefault
    'End If
End Sub

Private Sub Retirar_Click()

    If Not MainTimer.Check(TimersIndex.BuySell) Then
        Exit Sub
    End If
    
    If Val(Cantidad2.Text) <= 0 Or Val(Cantidad2.Text) > UserBankGold Or Not IsNumeric(Cantidad2.Text) Then
        Cantidad2.ForeColor = vbRed
        Exit Sub
    End If
                                                                    
    Cantidad2.ForeColor = &HC0FFFF
                                                            
    Call WriteBankExtractGold(Val(Cantidad2.Text))
                                                                    
    UserBankGold = UserBankGold - Val(Cantidad2.Text)
    lblUserBankGold.Caption = PonerPuntos(UserBankGold)

    UserGld = UserGld + Val(Cantidad2.Text)
    frmMain.GldLbl.Caption = PonerPuntos(UserGld)
    
    Call InitGld(Val(Cantidad2.Text))
End Sub
