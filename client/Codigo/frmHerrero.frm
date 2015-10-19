VERSION 5.00
Begin VB.Form frmHerrero 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "FrizQuadrata BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5175
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "1"
      Top             =   2940
      Width           =   1050
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ComboBox cboItemsCiclo 
      BackColor       =   &H80000006&
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
      Height          =   315
      Left            =   5325
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   3945
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   2355
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLingotes0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCantidadCiclo 
      Height          =   645
      Left            =   5160
      Top             =   3435
      Width           =   1110
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   4
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   3
      Left            =   1560
      Top             =   2985
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   2
      Left            =   1560
      Top             =   2190
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   2
      Left            =   5280
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   720
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   720
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   720
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image picMejorar1 
      Height          =   420
      Left            =   3360
      Top             =   2370
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image picMejorar2 
      Height          =   420
      Left            =   3360
      Top             =   3180
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   2760
      Top             =   4650
      Width           =   1455
   End
   Begin VB.Image picConstruir3 
      Height          =   420
      Left            =   3360
      Top             =   3960
      Width           =   1710
   End
   Begin VB.Image picConstruir2 
      Height          =   420
      Left            =   3360
      Top             =   3180
      Width           =   1710
   End
   Begin VB.Image picConstruir1 
      Height          =   420
      Left            =   3360
      Top             =   2370
      Width           =   1710
   End
   Begin VB.Image picCheckBox 
      Height          =   420
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   1860
      Width           =   435
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   2
      Left            =   3240
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image picMejorar3 
      Height          =   420
      Left            =   3360
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   1
      Left            =   1560
      Top             =   1395
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image picMejorar0 
      Height          =   420
      Left            =   3360
      Top             =   1560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image picConstruir0 
      Height          =   420
      Left            =   3360
      Top             =   1560
      Width           =   1710
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ePestania
    ieArmas
    ieArmaduras
    ieMejorar
End Enum

Private picCheck As Picture
Private picRecuadroItem As Picture
Private picRecuadroLingotes As Picture

Private Pestanias(0 To 2) As Picture
Private UltimaPestania As Byte

Private cPicCerrar As clsGraphicalButton
Private cPicConstruir(0 To 3) As clsGraphicalButton
Private cPicMejorar(0 To 3) As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private Cargando As Boolean

Private UsarMacro As Boolean
Private Armas As Boolean

Private clsFormulario As clsFormMovementManager

Private Sub CargarImagenes()
    Dim Index As Integer

    Set Pestanias(ePestania.ieArmas) = LoadPicture(GrhPath & "VentanaHerreríaArmas.jpg")
    Set Pestanias(ePestania.ieArmaduras) = LoadPicture(GrhPath & "VentanaHerreríaArmaduras.jpg")
    Set Pestanias(ePestania.ieMejorar) = LoadPicture(GrhPath & "VentanaHerreríaMejorar.jpg")
    
    Set picCheck = LoadPicture(GrhPath & "CheckBoxHerrería.jpg")
    
    Set picRecuadroItem = LoadPicture(GrhPath & "RecuadroItemsHerreria.jpg")
    'Set picRecuadroLingotes = LoadPicture(GrhPath & "RecuadroLingotes.jpg")
    
    For Index = 1 To MAX_LIST_Items
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoUpgrade(Index).Picture = picRecuadroItem
        imgMarcoLingotes(Index).Picture = picRecuadroLingotes
    Next Index
    
    Set cPicCerrar = New clsGraphicalButton
    Set cPicConstruir(0) = New clsGraphicalButton
    Set cPicConstruir(1) = New clsGraphicalButton
    Set cPicConstruir(2) = New clsGraphicalButton
    Set cPicConstruir(3) = New clsGraphicalButton
    Set cPicMejorar(0) = New clsGraphicalButton
    Set cPicMejorar(1) = New clsGraphicalButton
    Set cPicMejorar(2) = New clsGraphicalButton
    Set cPicMejorar(3) = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton
    
    Call cPicCerrar.Initialize(imgCerrar, GrhPath & "BotónCerrarHerrería.jpg", GrhPath & "BotónCerrarRolloverHerrería.jpg", GrhPath & "BotónCerrarClickHerrería.jpg", Me)
    Call cPicConstruir(0).Initialize(picConstruir0, GrhPath & "BotónConstruirHerreria.jpg", GrhPath & "BotónConstruirRolloverHerreria.jpg", GrhPath & "BotónConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(1).Initialize(picConstruir1, GrhPath & "BotónConstruirHerreria.jpg", GrhPath & "BotónConstruirRolloverHerreria.jpg", GrhPath & "BotónConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(2).Initialize(picConstruir2, GrhPath & "BotónConstruirHerreria.jpg", GrhPath & "BotónConstruirRolloverHerreria.jpg", GrhPath & "BotónConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(3).Initialize(picConstruir3, GrhPath & "BotónConstruirHerreria.jpg", GrhPath & "BotónConstruirRolloverHerreria.jpg", GrhPath & "BotónConstruirClickHerreria.jpg", Me)
    Call cPicMejorar(0).Initialize(picMejorar0, GrhPath & "BotónMejorarHerreria.jpg", GrhPath & "BotónMejorarRolloverHerreria.jpg", GrhPath & "BotónMejorarClickHerreria.jpg", Me)
    Call cPicMejorar(1).Initialize(picMejorar1, GrhPath & "BotónMejorarHerreria.jpg", GrhPath & "BotónMejorarRolloverHerreria.jpg", GrhPath & "BotónMejorarClickHerreria.jpg", Me)
    Call cPicMejorar(2).Initialize(picMejorar2, GrhPath & "BotónMejorarHerreria.jpg", GrhPath & "BotónMejorarRolloverHerreria.jpg", GrhPath & "BotónMejorarClickHerreria.jpg", Me)
    Call cPicMejorar(3).Initialize(picMejorar3, GrhPath & "BotónMejorarHerreria.jpg", GrhPath & "BotónMejorarRolloverHerreria.jpg", GrhPath & "BotónMejorarClickHerreria.jpg", Me)

    imgCantidadCiclo.Picture = LoadPicture(GrhPath & "ConstruirPorCiclo.jpg")
    
    picPestania(ePestania.ieArmas).MouseIcon = picMouseIcon
    picPestania(ePestania.ieArmaduras).MouseIcon = picMouseIcon
    picPestania(ePestania.ieMejorar).MouseIcon = picMouseIcon
    
    picCheckBox.MouseIcon = picMouseIcon
End Sub

Private Sub ConstruirItem(ByVal Index As Integer)

    If Not MainTimer.Check(TimersIndex.Work) Then
        Exit Sub
    End If

    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible Then ItemIndex = Scroll.Value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmasHerrero(ItemIndex).ObjIndex
                frmMain.ActivarMacroTrabajo
            Else
                'Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmasHerrero(ItemIndex).ObjIndex)
            
        Case ePestania.ieArmaduras
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmadurasHerrero(ItemIndex).ObjIndex
                frmMain.ActivarMacroTrabajo
             Else
                'Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmadurasHerrero(ItemIndex).ObjIndex)
        
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(HerreroMejorar(ItemIndex).ObjIndex)
    End Select
    
    Unload Me

End Sub

Private Sub Form_Load()
    Dim MaxConstItem As Integer
    Dim i As Integer
    
    'Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    CargarImagenes
    
    'Cargar imagenes
    Set Picture = Pestanias(ePestania.ieArmas)
    picCheckBox.Picture = picCheck
    
    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 4) / 5)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    
    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    
    cboItemsCiclo.ListIndex = 0
    
    Cargando = False
    
    UsarMacro = True
    Armas = True
    UltimaPestania = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
    Dim i As Integer
    
    picLingotes0.Visible = (NumItems >= 1)
    picLingotes1.Visible = (NumItems >= 2)
    picLingotes2.Visible = (NumItems >= 3)
    picLingotes3.Visible = (NumItems >= 4)
    
    For i = 1 To MAX_LIST_Items
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoLingotes(i).Visible = (NumItems >= i)
        picUpgradeItem(i).Visible = (NumItems >= i And Upgrading)
        imgMarcoUpgrade(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    picConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    picConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    picConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    picConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    picMejorar0.Visible = (NumItems >= 1 And Upgrading)
    picMejorar1.Visible = (NumItems >= 2 And Upgrading)
    picMejorar2.Visible = (NumItems >= 3 And Upgrading)
    picMejorar3.Visible = (NumItems >= 4 And Upgrading)
    
    picCheckBox.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    imgCantidadCiclo.Visible = Not Upgrading And UsarMacro
    txtCantItems.Visible = Not Upgrading
    picCheckBox.Visible = Not Upgrading
    
    If NumItems > MAX_LIST_Items Then
        Scroll.Visible = True
        Cargando = True
        Scroll.max = NumItems - MAX_LIST_Items
        Cargando = False
    Else
        Scroll.Visible = False
    End If
End Sub

Private Sub RenderItem(ByRef Pic As PictureBox, ByVal GrhIndex As Long)
    Dim SR As RECT
    Dim DR As RECT
    
    With GrhData(GrhIndex)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .PixelWidth
        SR.Bottom = SR.Top + .PixelHeight
    End With
    
    DR.Right = 32
    DR.Bottom = 32
    
    Call DrawGrhtoHdc(Pic.hdc, GrhIndex, SR, DR)
    Pic.Refresh
End Sub

Public Sub RenderList(ByVal Inicio As Integer, ByVal Armas As Boolean)
Dim i As Long
Dim NumItems As Integer
Dim ObjHerrero() As tItemsConstruibles

If Armas Then
    ObjHerrero = ArmasHerrero
Else
    ObjHerrero = ArmadurasHerrero
End If

NumItems = UBound(ObjHerrero)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_Items
    If i + Inicio <= NumItems Then
        With ObjHerrero(i + Inicio)
            'Agrego el Item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .Name
            
             'Inventariode lingotes
            Call InvLingosHerreria(i).SetSlot(1, 0, .LinH, LH_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro", True, False)
            Call InvLingosHerreria(i).SetSlot(2, 0, .LinP, LP_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Plata", True, False)
            Call InvLingosHerreria(i).SetSlot(3, 0, .LinO, LO_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Oro", True, False)
        End With
    End If
Next i
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(HerreroMejorar)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_Items
    If i + Inicio <= NumItems Then
        With HerreroMejorar(i + Inicio)
            'Agrego el Item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .Name
            
            Call RenderItem(picUpgradeItem(i), .UpgradeGrhIndex)
            picUpgradeItem(i).ToolTipText = .UpgradeName
            
             'Inventariode lingotes
            Call InvLingosHerreria(i).SetSlot(1, 0, .LinH, LH_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro", True, False)
            Call InvLingosHerreria(i).SetSlot(2, 0, .LinP, LP_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Plata", True, False)
            Call InvLingosHerreria(i).SetSlot(3, 0, .LinO, LO_Grh, 0, 0, 0, 0, 0, 0, "Lingotes de Oro", True, False)
        End With
    End If
Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub picCheckBox_Click()
    
    UsarMacro = Not UsarMacro

    If UsarMacro Then
        picCheckBox.Picture = picCheck
    Else
        picCheckBox.Picture = Nothing
    End If
    
    cboItemsCiclo.Visible = UsarMacro
    imgCantidadCiclo.Visible = UsarMacro
End Sub

Private Sub picConstruir0_Click()
    Call ConstruirItem(1)
End Sub

Private Sub picConstruir1_Click()
    Call ConstruirItem(2)
End Sub

Private Sub picConstruir2_Click()
    Call ConstruirItem(3)
End Sub

Private Sub picConstruir3_Click()
    Call ConstruirItem(4)
End Sub

Private Sub picLingotes0_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub picLingotes1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub picLingotes2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub picLingotes3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub picMejorar0_Click()
    Call ConstruirItem(1)
End Sub

Private Sub picMejorar1_Click()
    Call ConstruirItem(2)
End Sub

Private Sub picMejorar2_Click()
    Call ConstruirItem(3)
End Sub

Private Sub picMejorar3_Click()
    Call ConstruirItem(4)
End Sub

Private Sub picPestania_Click(Index As Integer)
    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then
        Exit Sub
    End If
    
    If UltimaPestania = Index Then
        Exit Sub
    End If
    
    Scroll.Value = 0
    
    Select Case Index
        Case ePestania.ieArmas
            'Background
            Picture = Pestanias(ePestania.ieArmas)
            
            NumItems = UBound(ArmasHerrero)
        
            Call HideExtraControls(NumItems)
            
            'Cargo inventarios e imagenes
            Call RenderList(1, True)
            
            Armas = True
            
        Case ePestania.ieArmaduras
            'Background
            Picture = Pestanias(ePestania.ieArmaduras)
            
            NumItems = UBound(ArmadurasHerrero)
        
            Call HideExtraControls(NumItems)
            
            'Cargo inventarios e imagenes
            Call RenderList(1, False)
            
            Armas = False
            
        Case ePestania.ieMejorar
            'Background
            Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(HerreroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)
    End Select

    UltimaPestania = Index
End Sub

Private Sub Scroll_Change()
    Dim i As Long
    
    If Cargando Then
        Exit Sub
    End If
    
    i = Scroll.Value
    'Cargo inventarios e imagenes
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
            Call RenderList(i + 1, True)
        Case ePestania.ieArmaduras
            Call RenderList(i + 1, False)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub

Private Sub txtCantItems_Change()
On Error GoTo ErrHandler
    If Val(txtCantItems.Text) < 0 Then
        txtCantItems.Text = 1
    End If
    
    If Val(txtCantItems.Text) > MaxInvObjs Then
        txtCantItems.Text = MaxInvObjs
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantItems.Text = MaxInvObjs
End Sub

Private Sub txtCantItems_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

