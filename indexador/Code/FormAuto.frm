VERSION 5.00
Begin VB.Form FormAuto 
   Caption         =   "Indexador automatico"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   ScaleHeight     =   5955
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ISimple 
      Caption         =   "simple"
      Height          =   255
      Left            =   5280
      TabIndex        =   62
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Animacion Avanzada"
      Height          =   255
      Left            =   2280
      TabIndex        =   44
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Animacion normal"
      Height          =   255
      Left            =   480
      TabIndex        =   43
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Animacion Especial"
      Height          =   5415
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Width           =   6255
      Begin VB.OptionButton Optiondimension 
         Caption         =   " (pj) alkon"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   61
         Top             =   2200
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FormAuto.frx":0000
         Left            =   1680
         List            =   "FormAuto.frx":000D
         TabIndex        =   60
         Text            =   "Combo2"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox CheckAuto 
         Caption         =   "Check1"
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   3840
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4440
         TabIndex        =   56
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   55
         Text            =   "1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.OptionButton Optiondimension 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   51
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Optiondimension 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   50
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   8
         Left            =   3720
         TabIndex        =   47
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   45
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox ComboTipoAnim 
         Height          =   315
         ItemData        =   "FormAuto.frx":0025
         Left            =   840
         List            =   "FormAuto.frx":002F
         TabIndex        =   42
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton CommandCalu2 
         Caption         =   "AutoCalcular"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton CommandBuscar2 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   31
         Text            =   "22"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   30
         Text            =   "4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   28
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   27
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   24
         Text            =   "6"
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Labelbody2 
         Caption         =   "offset:"
         Height          =   255
         Left            =   3840
         TabIndex        =   58
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Labelbody1 
         Caption         =   "i:"
         Height          =   255
         Left            =   2400
         TabIndex        =   57
         Top             =   3840
         Width           =   135
      End
      Begin VB.Label Labelbody 
         Caption         =   "Autoindexar"
         Height          =   255
         Left            =   720
         TabIndex        =   54
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Normal(NPC)"
         Height          =   255
         Left            =   4800
         TabIndex        =   53
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "(pj) clasico"
         Height          =   255
         Left            =   4800
         TabIndex        =   52
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Loffy 
         Caption         =   "y"
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   3840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Loffx 
         Caption         =   "x"
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Loff 
         Caption         =   "Offset:"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   3000
         X2              =   3360
         Y1              =   3240
         Y2              =   3000
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3000
         X2              =   3360
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Label Label3 
         Caption         =   "Totales"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   41
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "A lo alto"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "A lo ancho"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Ltext 
         Caption         =   "BMP:"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   38
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Ancho:"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   37
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Alto:"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   36
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Primer indice:"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   35
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Ltext 
         Caption         =   "nºframes:"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Indexacion simple"
      Height          =   5415
      Index           =   2
      Left            =   0
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command6 
         Caption         =   "indexar"
         Height          =   255
         Left            =   2160
         TabIndex        =   69
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "buscar"
         Height          =   255
         Left            =   3960
         TabIndex        =   68
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   66
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   65
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Numero Grh:"
         Height          =   255
         Left            =   600
         TabIndex        =   67
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Bmp:"
         Height          =   255
         Left            =   840
         TabIndex        =   64
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Animacion normal"
      Height          =   5415
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Crear"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   6
         Text            =   "1"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton CommandBuscar 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton CommandCalu 
         Caption         =   "AutoCalcular"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   2880
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   1
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Ltext 
         Caption         =   "nºframes:"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Primer indice:"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Ltext 
         Caption         =   "Alto:"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   20
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Ancho:"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   19
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "BMP:"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   18
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "A lo ancho"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "A lo alto"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Totales"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3000
         X2              =   3360
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3000
         X2              =   3360
         Y1              =   3240
         Y2              =   3000
      End
      Begin VB.Label Label4 
         Caption         =   "Mismo BMP"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "BMP consecutivos"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FormAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionNormales(1 To 22) As Position
Private PosicionNormales2(1 To 22) As Position

Private Sub CheckAuto_Click()
    If CheckAuto.value = vbChecked Then
        Text1.Enabled = True
        Text2.Enabled = True
    Else
        Text1.Enabled = False
        Text2.Enabled = False
    End If
End Sub



Private Sub ComboTipoAnim_Click()
        Combo2.Visible = False
        Labelbody.Visible = False
        Labelbody1.Visible = False
        Labelbody2.Visible = False
        Loff.Visible = False
        Loffx.Visible = False
        Loffy.Visible = False
        FormAuto.TextDatos2(7).Visible = False
        FormAuto.TextDatos2(8).Visible = False
        FormAuto.TextDatos2(0).Enabled = False
        FormAuto.TextDatos2(1).Enabled = False
        FormAuto.TextDatos2(6).Enabled = False
        Text1.Visible = False
        Text2.Visible = False
        CheckAuto.Visible = False
        Optiondimension(0).Visible = False
        Optiondimension(1).Visible = False
        Optiondimension(2).Visible = False
        Label5.Visible = False
        Label6.Visible = False
Select Case ComboTipoAnim.listIndex
    Case 0
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
        Labelbody.Visible = True
        Labelbody1.Visible = True
        Labelbody2.Visible = True
        Text1.Visible = True
        Text1.Enabled = False
        Text2.Visible = True
        Text2.Enabled = False
        CheckAuto.Visible = True
        CheckAuto.value = vbUnchecked
        Text1.Text = UBound(BodyData) + 1
        Text2.Text = "-38º0"
        Combo2.Visible = True
        Combo2.listIndex = 0
        Optiondimension(0).Visible = True
        Optiondimension(1).Visible = True
        Optiondimension(2).Visible = True
        Optiondimension(0).value = True
        Label5.Visible = True
        Label6.Visible = True
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
    Case 1
        Loff.Visible = True
        Loffx.Visible = True
        Loffy.Visible = True
        FormAuto.TextDatos2(7).Visible = True
        FormAuto.TextDatos2(8).Visible = True
        FormAuto.TextDatos2(0).Enabled = True
        FormAuto.TextDatos2(1).Enabled = True
        FormAuto.TextDatos2(6).Enabled = True
        FormAuto.TextDatos2(2).Enabled = True
        FormAuto.TextDatos2(3).Enabled = True
        FormAuto.CommandCalu2.Enabled = True
End Select
End Sub


Private Sub Command1_Click()
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim ActualFrame As Integer
Dim curX As Long, curY As Long

On Error GoTo errh

For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) > 32000 Or Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)


If (Not hayGrHlibres(PrimerIndice, FramesTotales + 1)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
Exit Sub
End If

ActualFrame = 0
curX = 0
curY = 0
If Option1(0).value Then
    ' Frames en el mismo BMP
    For FramesY = 1 To FramesAlto
        For FramesX = 1 To FramesAncho
            Grhdata(PrimerIndice + ActualFrame).FileNum = NumeroBMP
            Grhdata(PrimerIndice + ActualFrame).Frames(1) = PrimerIndice + ActualFrame
            Grhdata(PrimerIndice + ActualFrame).NumFrames = 1
            Grhdata(PrimerIndice + ActualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + ActualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + ActualFrame).sX = Ancho * curX
            Grhdata(PrimerIndice + ActualFrame).sY = Alto * curY
            Grhdata(PrimerIndice + ActualFrame).TileHeight = Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + ActualFrame).TileWidth = Grhdata(PrimerIndice + ActualFrame).pixelWidth / TilePixelWidth
            curX = curX + 1
            ActualFrame = ActualFrame + 1
            If ActualFrame >= FramesTotales Then GoTo TerminarAnim
        Next FramesX
        curX = 0
        curY = curY + 1
    Next FramesY
    
Else
    For FramesY = 1 To FramesTotales
            Grhdata(PrimerIndice + ActualFrame).FileNum = NumeroBMP + ActualFrame
            Grhdata(PrimerIndice + ActualFrame).Frames(1) = PrimerIndice + ActualFrame
            Grhdata(PrimerIndice + ActualFrame).NumFrames = 1
            Grhdata(PrimerIndice + ActualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + ActualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + ActualFrame).sX = 0
            Grhdata(PrimerIndice + ActualFrame).sY = 0
            Grhdata(PrimerIndice + ActualFrame).TileHeight = Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + ActualFrame).TileWidth = Grhdata(PrimerIndice + ActualFrame).pixelWidth / TilePixelWidth
            ActualFrame = ActualFrame + 1
            If ActualFrame >= FramesTotales Then GoTo TerminarAnim
    Next FramesY
End If


TerminarAnim:

EstadoNoGuardado(e_EstadoIndexador.Grh) = True
Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesTotales

Grhdata(PrimerIndice + FramesTotales).pixelHeight = Grhdata(PrimerIndice).pixelHeight
Grhdata(PrimerIndice + FramesTotales).pixelWidth = Grhdata(PrimerIndice).pixelWidth
Grhdata(PrimerIndice + FramesTotales).sX = Grhdata(PrimerIndice).sX
Grhdata(PrimerIndice + FramesTotales).sY = Grhdata(PrimerIndice).sY
Grhdata(PrimerIndice + FramesTotales).TileHeight = Grhdata(PrimerIndice).TileHeight
Grhdata(PrimerIndice + FramesTotales).TileWidth = Grhdata(PrimerIndice).TileWidth
            

For ii = 1 To FramesTotales
    Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
Next ii
Grhdata(PrimerIndice + FramesTotales).Speed = 1

Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
Call frmMain.BuscarNuevoF(PrimerIndice)

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number

End Sub

Private Sub Command2_Click()
' creacion de un grafico normal:

Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim ActualFrame As Integer
Dim curX As Long, curY As Long
Dim respuesta As Byte

On Error GoTo errh

'Comprobamos si hay datos invalidos:
For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) > 32000 Or Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii

'Recogemos los valores necesarios
FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

'comprobamos que hay hueco
If (Not hayGrHlibres(PrimerIndice, FramesTotales + 4)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
    
    Exit Sub 'Realmente el implementar el sobreescribir seria solo quitar esta linea. Pero , sin la opcion de deshacer ahora mismo, no es muy recomendable
End If

If CheckAuto.value = vbChecked Then
    ii = Val(Text1.Text)
    If errorEnIndice() Then
        respuesta = MsgBox("El grafico Indice de autoindexacion indicado ya existe, ¿estas seguro de sobreescribirlo?", 4, "Aviso")
    Else
        respuesta = vbYes
    End If
    If respuesta <> vbYes Then Exit Sub
End If
ActualFrame = 0
curX = 0
curY = 0
If ComboTipoAnim.listIndex = 0 Then
    If Optiondimension(0).value Then
        ' Frames en el mismo BMP
        For FramesY = 1 To FramesTotales
                Grhdata(PrimerIndice + FramesY - 1).FileNum = NumeroBMP
                Grhdata(PrimerIndice + FramesY - 1).Frames(1) = FramesY
                Grhdata(PrimerIndice + FramesY - 1).NumFrames = 1
                Grhdata(PrimerIndice + FramesY - 1).pixelHeight = Alto
                Grhdata(PrimerIndice + FramesY - 1).pixelWidth = Ancho
                Grhdata(PrimerIndice + FramesY - 1).sX = PosicionNormales(FramesY).X
                Grhdata(PrimerIndice + FramesY - 1).sY = PosicionNormales(FramesY).Y
                Grhdata(PrimerIndice + FramesY - 1).TileHeight = Grhdata(PrimerIndice + FramesY - 1).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + FramesY - 1).TileWidth = Grhdata(PrimerIndice + FramesY - 1).pixelWidth / TilePixelWidth
        Next FramesY
        
    ElseIf Optiondimension(1).value Then
        For FramesY = 1 To 4
            For FramesX = 1 To FramesAncho
                Grhdata(PrimerIndice + ActualFrame).FileNum = NumeroBMP
                Grhdata(PrimerIndice + ActualFrame).Frames(1) = PrimerIndice + ActualFrame
                Grhdata(PrimerIndice + ActualFrame).NumFrames = 1
                Grhdata(PrimerIndice + ActualFrame).pixelHeight = Alto
                Grhdata(PrimerIndice + ActualFrame).pixelWidth = Ancho
                Grhdata(PrimerIndice + ActualFrame).sX = Ancho * curX
                Grhdata(PrimerIndice + ActualFrame).sY = Alto * curY
                Grhdata(PrimerIndice + ActualFrame).TileHeight = Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + ActualFrame).TileWidth = Grhdata(PrimerIndice + ActualFrame).pixelWidth / TilePixelWidth
                curX = curX + 1
                ActualFrame = ActualFrame + 1
                If ActualFrame >= FramesTotales Then GoTo TerminarAnim
            Next FramesX
            curX = 0
            curY = curY + 1
        Next FramesY
    Else
        For FramesY = 1 To FramesTotales
                Grhdata(PrimerIndice + FramesY - 1).FileNum = NumeroBMP
                Grhdata(PrimerIndice + FramesY - 1).Frames(1) = FramesY
                Grhdata(PrimerIndice + FramesY - 1).NumFrames = 1
                Grhdata(PrimerIndice + FramesY - 1).pixelHeight = Alto
                Grhdata(PrimerIndice + FramesY - 1).pixelWidth = Ancho
                Grhdata(PrimerIndice + FramesY - 1).sX = PosicionNormales2(FramesY).X
                Grhdata(PrimerIndice + FramesY - 1).sY = PosicionNormales2(FramesY).Y
                Grhdata(PrimerIndice + FramesY - 1).TileHeight = Grhdata(PrimerIndice + FramesY - 1).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + FramesY - 1).TileWidth = Grhdata(PrimerIndice + FramesY - 1).pixelWidth / TilePixelWidth
        Next FramesY
    End If
ElseIf ComboTipoAnim.listIndex = 1 Then
    For FramesY = 1 To FramesAlto
        For FramesX = 1 To FramesAncho
            Grhdata(PrimerIndice + ActualFrame).FileNum = NumeroBMP
            Grhdata(PrimerIndice + ActualFrame).Frames(1) = PrimerIndice + ActualFrame
            Grhdata(PrimerIndice + ActualFrame).NumFrames = 1
            Grhdata(PrimerIndice + ActualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + ActualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + ActualFrame).sX = Ancho * curX + Val(TextDatos2(7).Text)
            Grhdata(PrimerIndice + ActualFrame).sY = Alto * curY + Val(TextDatos2(8).Text)
            Grhdata(PrimerIndice + ActualFrame).TileHeight = Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + ActualFrame).TileWidth = Grhdata(PrimerIndice + ActualFrame).pixelWidth / TilePixelWidth
            curX = curX + 1
            ActualFrame = ActualFrame + 1
            If ActualFrame >= FramesTotales Then GoTo TerminarAnim
        Next FramesX
        curX = 0
        curY = curY + 1
    Next FramesY
End If

    

' No me gustan los Goto pero... es lo q hay xD
TerminarAnim:
 EstadoNoGuardado(e_EstadoIndexador.Grh) = True
If ComboTipoAnim.listIndex = 0 Then
    If Optiondimension(0).value Then 'indexacion clasica de bodys
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales).Speed = 1
        For ii = 1 To 6
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 1
        For ii = 7 To 12
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - 6) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 1
        For ii = 13 To 17
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - 12) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 1
        For ii = 18 To 22
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - 17) = PrimerIndice + ii - 1
        Next ii
    ElseIf Optiondimension(1).value Then 'indexacion npc standar
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesAncho
        Grhdata(PrimerIndice + FramesTotales).Speed = 1
        For ii = 1 To FramesAncho
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = FramesAncho
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 1
        For ii = FramesAncho + 1 To FramesAncho * 2
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - FramesAncho) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = FramesAncho
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 1
        For ii = (FramesAncho * 2) + 1 To FramesAncho * 3
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - (FramesAncho * 2)) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = FramesAncho
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 1
        For ii = (FramesAncho * 3) + 1 To FramesAncho * 4
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - (FramesAncho * 3)) = PrimerIndice + ii - 1
        Next ii
    Else 'indexacion ultimos bodys alkon
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales).Speed = 1
        For ii = 1 To 6
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 1
        For ii = 7 To 12
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - 6) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 1
        For ii = 13 To 17
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - 12) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 1
        For ii = 18 To 22
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - 17) = PrimerIndice + ii - 1
        Next ii
    End If
ElseIf ComboTipoAnim.listIndex = 1 Then ' animacion con offset
    Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
    Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesTotales
    Grhdata(PrimerIndice + FramesTotales).pixelHeight = Grhdata(PrimerIndice).pixelHeight
    Grhdata(PrimerIndice + FramesTotales).pixelWidth = Grhdata(PrimerIndice).pixelWidth
    Grhdata(PrimerIndice + FramesTotales).sX = Grhdata(PrimerIndice).sX
    Grhdata(PrimerIndice + FramesTotales).sY = Grhdata(PrimerIndice).sY
    Grhdata(PrimerIndice + FramesTotales).TileHeight = Grhdata(PrimerIndice).TileHeight
    Grhdata(PrimerIndice + FramesTotales).TileWidth = Grhdata(PrimerIndice).TileWidth
    
    For ii = 1 To FramesTotales
        Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
    Next ii
    Grhdata(PrimerIndice + FramesTotales).Speed = 1
End If

If CheckAuto.value = vbChecked Then ' selecionado autoindexar como...
        ii = Val(Text1.Text)
        Select Case Combo2.listIndex
            Case 0 'body
                Call AgregaBody(ii, False)
                BodyData(ii).HeadOffset.Y = Val(ReadField(1, Text2.Text, Asc("º")))
                BodyData(ii).HeadOffset.X = Val(ReadField(2, Text2.Text, Asc("º")))
                BodyData(ii).Walk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                BodyData(ii).Walk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                BodyData(ii).Walk(3).GrhIndex = PrimerIndice + FramesTotales
                BodyData(ii).Walk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Body) = True
            Case 1 'arma
                Call AgregaArma(ii, False)
                WeaponAnimData(ii).WeaponWalk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                WeaponAnimData(ii).WeaponWalk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                WeaponAnimData(ii).WeaponWalk(3).GrhIndex = PrimerIndice + FramesTotales
                WeaponAnimData(ii).WeaponWalk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Armas) = True
            Case 2 'escudo
                Call AgregaEscudo(ii, False)
                ShieldAnimData(ii).ShieldWalk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                ShieldAnimData(ii).ShieldWalk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                ShieldAnimData(ii).ShieldWalk(3).GrhIndex = PrimerIndice + FramesTotales
                ShieldAnimData(ii).ShieldWalk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Escudos) = True
            Case 3 'botas
                Call AgregaBota(ii, False)
                BotasAnimData(ii).Head(1).GrhIndex = PrimerIndice + FramesTotales + 1
                BotasAnimData(ii).Head(2).GrhIndex = PrimerIndice + FramesTotales + 3
                BotasAnimData(ii).Head(3).GrhIndex = PrimerIndice + FramesTotales
                BotasAnimData(ii).Head(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Botas) = True
            Case 4 'alas
                Call AgregaCapa(ii, False)
                EspaldaAnimData(ii).Head(1).GrhIndex = PrimerIndice + FramesTotales + 1
                EspaldaAnimData(ii).Head(2).GrhIndex = PrimerIndice + FramesTotales + 3
                EspaldaAnimData(ii).Head(3).GrhIndex = PrimerIndice + FramesTotales
                EspaldaAnimData(ii).Head(4).GrhIndex = PrimerIndice + FramesTotales + 2
                EstadoNoGuardado(e_EstadoIndexador.Cascos) = True
        End Select
End If

Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
Call frmMain.BuscarNuevoF(PrimerIndice)

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number

End Sub

Private Sub Command3_Click()
FormAuto.FrameAnim(0).Visible = True
FormAuto.FrameAnim(1).Visible = False
FormAuto.FrameAnim(2).Visible = False
End Sub

Private Sub Command4_Click()
FormAuto.FrameAnim(1).Visible = True
FormAuto.FrameAnim(0).Visible = False
FormAuto.FrameAnim(2).Visible = False
End Sub

Private Sub Command5_Click()
FormAuto.TextDatos3(5).Text = BuscarGrHlibres(1)
End Sub

Private Sub Command6_Click()
On Error GoTo errh
Dim FramesTotales As Integer
Dim respuesta As Byte
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

For ii = 4 To 5
    If Val(FormAuto.TextDatos3(ii).Text) > 32000 Or Val(FormAuto.TextDatos3(ii).Text) <= 0 Then
        FormAuto.TextDatos3(ii).Text = 0
    End If
Next ii



PrimerIndice = Val(FormAuto.TextDatos3(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text)



If (Not hayGrHlibres(PrimerIndice, 1)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    respuesta = MsgBox("El grafico Indice de autoindexacion indicado ya existe, ¿estas seguro de sobreescribirlo?", 4, "Aviso")
    If respuesta <> vbYes Then Exit Sub ' al ser solo un grafico permitimos sobreescribir ya que no conlleva el mucho riesgo( como mucho 1 grafico ^^)
    
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    Grhdata(PrimerIndice).FileNum = NumeroBMP
    Grhdata(PrimerIndice).Frames(1) = PrimerIndice
    Grhdata(PrimerIndice).NumFrames = 1
    Grhdata(PrimerIndice).pixelHeight = AltoBMP
    Grhdata(PrimerIndice).pixelWidth = AnchoBMP
    Grhdata(PrimerIndice).sX = 0
    Grhdata(PrimerIndice).sY = 0
    Grhdata(PrimerIndice).TileHeight = Grhdata(PrimerIndice).pixelHeight / TilePixelHeight
    Grhdata(PrimerIndice).TileWidth = Grhdata(PrimerIndice).pixelWidth / TilePixelWidth
    ' Hay cambios!
    EstadoNoGuardado(e_EstadoIndexador.Grh) = True
Else
    MsgBox "El bmp " & NumeroBMP & " no existe."
    Exit Sub
End If

Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
Call frmMain.BuscarNuevoF(PrimerIndice)

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub CommandBuscar2_Click()
    FormAuto.TextDatos2(5).Text = BuscarGrHlibres(Val(FormAuto.TextDatos2(6).Text) + 4)
End Sub

Private Sub CommandBuscar_Click()
    FormAuto.TextDatos(5).Text = BuscarGrHlibres(Val(FormAuto.TextDatos(6).Text) + 1)
End Sub

Private Sub CommandCalu_Click()
On Error GoTo errh
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) > 32000 Or Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii
If Val(FormAuto.TextDatos(6).Text) > 25 Then
    FormAuto.TextDatos(6).Text = 25
ElseIf Val(FormAuto.TextDatos(6).Text) <= 0 Then
    FormAuto.TextDatos(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)

If FramesAncho < 1 Then Exit Sub
If FramesAlto < 1 Then Exit Sub

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    FormAuto.TextDatos(2).Text = CInt(AltoBMP / FramesAlto)
    FormAuto.TextDatos(3).Text = CInt(AnchoBMP / FramesAncho)
End If
Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub CommandCalu2_Click()
'Calcula automaticamente el ancho de los frames a partir de su numero
On Error GoTo errh
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) > 32000 Or Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii
If Val(FormAuto.TextDatos2(6).Text) > 25 And ComboTipoAnim.listIndex > 1 Then
    FormAuto.TextDatos2(6).Text = 25
ElseIf Val(FormAuto.TextDatos2(6).Text) <= 0 Then
    FormAuto.TextDatos2(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

If FramesAncho < 1 Then Exit Sub
If FramesAlto < 1 Then Exit Sub

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    FormAuto.TextDatos2(2).Text = CInt(AltoBMP / FramesAlto)
    FormAuto.TextDatos2(3).Text = CInt(AnchoBMP / FramesAncho)
End If

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub Form_Load()
Dim ii As Long
FormAuto.FrameAnim(0).Visible = True
FormAuto.FrameAnim(1).Visible = False
FormAuto.ComboTipoAnim.listIndex = 0
DibujarIndexaciones.activo = True
Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)

For ii = 1 To 6
    PosicionNormales(ii).Y = 0
Next ii

For ii = 7 To 12
    PosicionNormales(ii).Y = 45
Next ii

For ii = 13 To 17
    PosicionNormales(ii).Y = 90
Next ii

For ii = 18 To 22
    PosicionNormales(ii).Y = 135
Next ii

PosicionNormales(1).X = 0
PosicionNormales(2).X = 25
PosicionNormales(3).X = 49
PosicionNormales(4).X = 73
PosicionNormales(5).X = 98
PosicionNormales(6).X = 123
PosicionNormales(7).X = 0
PosicionNormales(8).X = 25
PosicionNormales(9).X = 49
PosicionNormales(10).X = 73
PosicionNormales(11).X = 98
PosicionNormales(12).X = 123
PosicionNormales(13).X = 0
PosicionNormales(14).X = 25
PosicionNormales(15).X = 49
PosicionNormales(16).X = 73
PosicionNormales(17).X = 98
PosicionNormales(18).X = 0
PosicionNormales(19).X = 25
PosicionNormales(20).X = 49
PosicionNormales(21).X = 73
PosicionNormales(22).X = 98


For ii = 1 To 6
    PosicionNormales2(ii).Y = 0
Next ii

For ii = 7 To 12
    PosicionNormales2(ii).Y = 45
Next ii

For ii = 13 To 17
    PosicionNormales2(ii).Y = 90
Next ii

For ii = 18 To 22
    PosicionNormales2(ii).Y = 135
Next ii

PosicionNormales2(1).X = 0
PosicionNormales2(2).X = 25
PosicionNormales2(3).X = 50
PosicionNormales2(4).X = 75
PosicionNormales2(5).X = 100
PosicionNormales2(6).X = 125
PosicionNormales2(7).X = 0
PosicionNormales2(8).X = 25
PosicionNormales2(9).X = 50
PosicionNormales2(10).X = 75
PosicionNormales2(11).X = 100
PosicionNormales2(12).X = 125
PosicionNormales2(13).X = 0
PosicionNormales2(14).X = 25
PosicionNormales2(15).X = 50
PosicionNormales2(16).X = 75
PosicionNormales2(17).X = 100
PosicionNormales2(18).X = 0
PosicionNormales2(19).X = 25
PosicionNormales2(20).X = 50
PosicionNormales2(21).X = 75
PosicionNormales2(22).X = 100


TextDatos(6).Text = 1
TextDatos(2).Text = 16
TextDatos(3).Text = 16
TextDatos(4).Text = 1
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DibujarIndexaciones.activo = False
End Sub

Private Sub ISimple_Click()
FormAuto.FrameAnim(0).Visible = False
FormAuto.FrameAnim(1).Visible = False
FormAuto.FrameAnim(2).Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
       TextDatos(0).Enabled = True
       TextDatos(1).Enabled = True
    Else
        TextDatos(0).Enabled = False
        TextDatos(1).Enabled = False
        TextDatos(0).Text = 1
        TextDatos(1).Text = 1
    End If
End Sub

Private Sub Optiondimension_Click(Index As Integer)
    
    Select Case Index
    Case 0
        CommandCalu2.Enabled = False
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
        TextDatos2(0).Enabled = False
        Combo2.Enabled = True
        Text2.Text = "-38º0"
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
    Case 1
        TextDatos2(0).Enabled = True
        FormAuto.TextDatos2(2).Enabled = True
        FormAuto.TextDatos2(3).Enabled = True
        Combo2.Enabled = False
        Combo2.listIndex = 0
        FormAuto.TextDatos2(1).Text = 0
        FormAuto.TextDatos2(1).Text = 4
        Text2.Text = "0º0"
    Case 2
        CommandCalu2.Enabled = False
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
        TextDatos2(0).Enabled = False
        Combo2.Enabled = True
        Text2.Text = "-38º0"
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 45
        FormAuto.TextDatos2(3).Text = 25
    End Select
    TextDatos2(0).Text = TextDatos2(0).Text
End Sub

Private Sub Text1_Change()
    On Error GoTo errh
    If Val(Text1.Text < 1) Then Text1.Text = 1
    If Val(Text1.Text > 32000) Then Text1.Text = 32000
    
    If errorEnIndice() Then
         Text1.BackColor = vbRed
    Else
        Text1.BackColor = vbWhite
    End If
    
    Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Function errorEnIndice() As Boolean
Select Case Combo2.listIndex
        Case 0 'body
            If Val(Text1.Text) <= UBound(BodyData) Then
                If BodyData(Val(Text1.Text)).Walk(1).GrhIndex > 0 Then
                    errorEnIndice = True
                End If
            End If
        Case 1 'arma
            If Val(Text1.Text) <= UBound(WeaponAnimData) Then
                If WeaponAnimData(Val(Text1.Text)).WeaponWalk(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
        Case 2 'escudo
            If Val(Text1.Text) <= UBound(ShieldAnimData) Then
                If ShieldAnimData(Val(Text1.Text)).ShieldWalk(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
        Case 3 'botas
            If Val(Text1.Text) <= UBound(BotasAnimData) Then
                Text1.BackColor = vbWhite
            Else
                If BotasAnimData(Val(Text1.Text)).Head(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
        Case 4 'alas
            If Val(Text1.Text) <= UBound(EspaldaAnimData) Then
                If EspaldaAnimData(Val(Text1.Text)).Head(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
    End Select
End Function
Private Sub Text2_Change()
Dim tempdouble1 As Double, tempdobule2 As Double

tempdouble1 = Val(ReadField(1, Text2.Text, Asc("º")))
tempdobule2 = Val(ReadField(2, Text2.Text, Asc("º")))

If tempdouble1 < -32000 Or tempdouble1 > 32000 Then
    Text2.Text = "0º" & tempdobule2
    tempdouble1 = 0
End If

If tempdobule2 < -32000 Or tempdobule2 > 32000 Then
    Text2.Text = tempdouble1 & "º0"
End If
        
End Sub

Private Sub TextDatos_Change(Index As Integer)
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
On Error GoTo errh
For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) > 32000 Or Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii
If Val(FormAuto.TextDatos(6).Text) > 25 Then
    FormAuto.TextDatos(6).Text = 25
ElseIf Val(FormAuto.TextDatos(6).Text) <= 0 Then
    FormAuto.TextDatos(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)

If FramesTotales > 0 Then
    FormAuto.CommandBuscar.Enabled = True
Else
    FormAuto.CommandBuscar.Enabled = False
End If

If Not hayGrHlibres(PrimerIndice, FramesTotales + 1) Then
    FormAuto.TextDatos(5).BackColor = vbRed
Else
    FormAuto.TextDatos(5).BackColor = vbWhite
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    FormAuto.TextDatos(4).BackColor = vbWhite
    CommandCalu.Enabled = True
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    Call CrearIndexacion(0)
    Call frmMain.BuscarNuevoF(NumeroBMP)
Else
    CommandCalu.Enabled = False
    FormAuto.TextDatos(4).BackColor = vbRed
    Exit Sub
End If

Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)

If FramesAlto * Alto > AltoBMP Then
     FormAuto.TextDatos(2).BackColor = vbYellow
Else
     FormAuto.TextDatos(2).BackColor = vbWhite
End If

If FramesAncho * Ancho > AnchoBMP Then
     FormAuto.TextDatos(3).BackColor = vbYellow
Else
     FormAuto.TextDatos(3).BackColor = vbWhite
End If

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub TextDatos2_Change(Index As Integer)
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Integer
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
On Error GoTo errh
For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) > 32000 Or Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii

If Val(FormAuto.TextDatos2(6).Text) > 25 And ComboTipoAnim.listIndex > 1 Then
    FormAuto.TextDatos2(6).Text = 25
ElseIf Val(FormAuto.TextDatos2(6).Text) <= 0 Then
    FormAuto.TextDatos2(6).Text = 0
End If


FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

If ComboTipoAnim.listIndex = 0 And Optiondimension(1).value Then
    FormAuto.TextDatos2(6).Text = FramesAncho * 4
End If

If FramesTotales > 0 Then
    FormAuto.CommandBuscar2.Enabled = True
Else
    FormAuto.CommandBuscar2.Enabled = False
End If

If Not hayGrHlibres(PrimerIndice, FramesTotales + 4) Then
    FormAuto.TextDatos2(5).BackColor = vbRed
Else
    FormAuto.TextDatos2(5).BackColor = vbWhite
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    FormAuto.TextDatos2(4).BackColor = vbWhite
    CommandCalu2.Enabled = (Optiondimension(1).value) Or ComboTipoAnim.listIndex = 1
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    Call CrearIndexacion(1)
    Call frmMain.BuscarNuevoF(NumeroBMP)
Else
    CommandCalu2.Enabled = False
    FormAuto.TextDatos2(4).BackColor = vbRed
    Exit Sub
End If

Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)

If FramesAlto * Alto > AltoBMP Then
     FormAuto.TextDatos2(2).BackColor = vbYellow
Else
     FormAuto.TextDatos2(2).BackColor = vbWhite
End If

If FramesAncho * Ancho > AnchoBMP Then
     FormAuto.TextDatos2(3).BackColor = vbYellow
Else
     FormAuto.TextDatos2(3).BackColor = vbWhite
End If
Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Public Sub CrearIndexacion(ByVal Index As Integer)
On Error Resume Next
Dim FramesY As Integer
Dim FramesX As Integer
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, Alto As Long, Ancho As Long
Dim ActualFrame As Integer
Dim curX As Integer, curY As Integer
Dim BitCount As Integer





If Index = 0 Then
    FramesTotales = Val(FormAuto.TextDatos(6).Text)
    NumeroBMP = Val(FormAuto.TextDatos(4).Text)
    Alto = Val(FormAuto.TextDatos(2).Text)
    Ancho = Val(FormAuto.TextDatos(3).Text)
    FramesAncho = Val(FormAuto.TextDatos(0).Text)
    FramesAlto = Val(FormAuto.TextDatos(1).Text)
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    DibujarIndexaciones.activo = True
    If Option1(0).value Then
        For FramesY = 1 To FramesAlto
            For FramesX = 1 To FramesAncho
                DibujarIndexaciones.Inicios(ActualFrame + 1).X = Ancho * curX
                DibujarIndexaciones.Inicios(ActualFrame + 1).Y = Alto * curY
                curX = curX + 1
                ActualFrame = ActualFrame + 1
            Next FramesX
            curX = 0
            curY = curY + 1
        Next FramesY
        DibujarIndexaciones.Total = FramesTotales
    Else
        DibujarIndexaciones.Inicios(1).X = 0
        DibujarIndexaciones.Inicios(1).Y = 0
        DibujarIndexaciones.Total = 1
    End If
ElseIf Index = 1 Then
    FramesTotales = Val(FormAuto.TextDatos2(6).Text)
    NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
    Alto = Val(FormAuto.TextDatos2(2).Text)
    Ancho = Val(FormAuto.TextDatos2(3).Text)
    FramesAncho = Val(FormAuto.TextDatos2(0).Text)
    FramesAlto = Val(FormAuto.TextDatos2(1).Text)
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    
    If ComboTipoAnim.listIndex = 0 Then
        If Optiondimension(0).value Then
            ' Frames en el mismo BMP
            DibujarIndexaciones.activo = True
            For FramesY = 1 To FramesTotales
                DibujarIndexaciones.Inicios(FramesY).X = PosicionNormales(FramesY).X
                DibujarIndexaciones.Inicios(FramesY).Y = PosicionNormales(FramesY).Y
            Next FramesY
            DibujarIndexaciones.Total = FramesTotales
        ElseIf Optiondimension(1).value Then
            DibujarIndexaciones.activo = True
            For FramesY = 1 To 4
                For FramesX = 1 To FramesAncho
                    DibujarIndexaciones.Inicios(ActualFrame + 1).X = Ancho * curX
                    DibujarIndexaciones.Inicios(ActualFrame + 1).Y = Alto * curY
                    curX = curX + 1
                    ActualFrame = ActualFrame + 1
                Next FramesX
                curX = 0
                curY = curY + 1
            Next FramesY
            DibujarIndexaciones.Total = FramesTotales
        Else
            DibujarIndexaciones.activo = True
            For FramesY = 1 To FramesTotales
                DibujarIndexaciones.Inicios(FramesY).X = PosicionNormales2(FramesY).X
                DibujarIndexaciones.Inicios(FramesY).Y = PosicionNormales2(FramesY).Y
            Next FramesY
            DibujarIndexaciones.Total = FramesTotales
        End If
    ElseIf ComboTipoAnim.listIndex = 1 Then
        DibujarIndexaciones.activo = True
        For FramesY = 1 To FramesAlto
            For FramesX = 1 To FramesAncho
                DibujarIndexaciones.Inicios(ActualFrame + 1).X = Ancho * curX + Val(TextDatos2(7).Text)
                DibujarIndexaciones.Inicios(ActualFrame + 1).Y = Alto * curY + Val(TextDatos2(8).Text)
                curX = curX + 1
                ActualFrame = ActualFrame + 1
            Next FramesX
            curX = 0
            curY = curY + 1
        Next FramesY
        DibujarIndexaciones.Total = FramesTotales
    End If
ElseIf Index = 3 Then
    FramesTotales = 1
    NumeroBMP = Val(FormAuto.TextDatos3(4).Text)
    If ExisteBMP(NumeroBMP) > 0 Then
        Call GetTamañoBMP(NumeroBMP, Alto, Ancho, BitCount)
    End If
    FramesAncho = 1
    FramesAlto = 1
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    DibujarIndexaciones.activo = True
    DibujarIndexaciones.Inicios(1).X = 0
    DibujarIndexaciones.Inicios(1).Y = 0
    DibujarIndexaciones.Total = 1
End If

End Sub

Private Sub TextDatos3_Change(Index As Integer)
Dim NumeroBMP As Long
Dim PrimerIndice As Integer
Dim ii As Long

For ii = 4 To 5
    If Val(FormAuto.TextDatos3(ii).Text) > 32000 Or Val(FormAuto.TextDatos3(ii).Text) <= 0 Then
        FormAuto.TextDatos3(ii).Text = 0
    End If
Next ii

PrimerIndice = Val(FormAuto.TextDatos3(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text)

If Not hayGrHlibres(PrimerIndice, 1) Then
    FormAuto.TextDatos3(5).BackColor = vbRed
Else
    FormAuto.TextDatos3(5).BackColor = vbWhite
End If

If Index = 4 Then
    If ExisteBMP(NumeroBMP) > 0 Then
        If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
        Call CrearIndexacion(3)
        Call frmMain.BuscarNuevoF(NumeroBMP)
        FormAuto.TextDatos3(4).BackColor = vbWhite
    Else
        FormAuto.TextDatos3(4).BackColor = vbRed
    End If
End If
    
End Sub
