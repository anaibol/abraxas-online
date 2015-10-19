VERSION 5.00
Begin VB.Form FrmSearch 
   Caption         =   "Busqueda"
   ClientHeight    =   3660
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotonSalir 
      Caption         =   "Salir"
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ListBox ListRepetidos 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label LGrhnum 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Menu mnuopcionesI 
      Caption         =   "Indexar como..."
      Begin VB.Menu Ibody 
         Caption         =   "Como body"
      End
      Begin VB.Menu IAnimacion 
         Caption         =   "Como Animacion"
      End
      Begin VB.Menu ICompleto 
         Caption         =   "Como grafico independiente"
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Bfin As Boolean
Option Explicit
Public Sub HacerBusquedaR()
Dim i As Long
Dim ii As Long
Bfin = False
For i = 1 To MAXGrH
    If GrhData(i).NumFrames > 0 Then
        For ii = i + 1 To MAXGrH
            If Bfin Then Exit Sub
            If GrhData(i).NumFrames = GrhData(ii).NumFrames And Not GrhData(i).NumFrames = 1 Then
                If GrhData(i).FileNum = GrhData(ii).FileNum Then
                    If GrhData(i).Frames(2) = GrhData(ii).Frames(2) Then
                        ListRepetidos.AddItem "Grh: " & i & " es igual al grh: " & ii
                    End If
                End If
            Else
                If GrhData(i).sX = GrhData(ii).sX And GrhData(i).sY = GrhData(ii).sY Then
                    Debug.Print GrhData(i).FileNum & ";" & GrhData(ii).FileNum
                    If GrhData(i).FileNum = GrhData(ii).FileNum Then
                        If GrhData(i).pixelWidth = GrhData(ii).pixelWidth Then
                            If GrhData(i).pixelHeight = GrhData(ii).pixelHeight Then
                                ListRepetidos.AddItem "Grh: " & i & " es igual al grh: " & ii
                            End If
                        End If
                    End If
                End If
            End If
        Next ii
    End If
    DoEvents
    LGrhnum.Caption = i
Next i
End Sub

Public Sub HacerBusquedaNI()
Dim i As Long
Dim ii As Long
Dim indexado As Boolean
Dim ExisteBMPActual As Byte
Dim UltimoGrafiCoRevisar As Long
Dim stringGrap As String

UltimoGrafiCoRevisar = 32000


Bfin = False
For ii = 1 To UltimoGrafiCoRevisar
    stringGrap = vbNullString
    ExisteBMPActual = ExisteBMP(ii)
    If ExisteBMPActual = ResourceFile Or (ResourceFile = 3 And ExisteBMPActual > 0) Then
        indexado = False
        For i = 1 To MAXGrH
            If Bfin Then Exit Sub
            If GrhData(i).NumFrames > 0 Then
                If GrhData(i).FileNum = ii Then
                    LGrhnum.Caption = ii & ".bmp indexado en grh: " & i
                    indexado = True
                    Exit For
                End If
            End If
         Next i
         If Not indexado Then
            If ResourceFile = 3 And ExisteBMPActual = 1 Then
                If ii > 0 And ii <= UBound(ResourceF.graficos) Then
                    If ResourceF.graficos(ii).tamaño > 0 Then
                        stringGrap = " + ResF"
                    End If
                End If
            End If
            ListRepetidos.AddItem ii & ".bmp NO esta indexado. El grafico esta en : " & StringRecurso(ExisteBMPActual) & stringGrap
         End If
    End If
    LGrhnum.Caption = "bmp->" & ii
    DoEvents
Next ii
End Sub
Private Sub BotonSalir_Click()
Bfin = True
Unload Me
End Sub
Private Sub Form_close()
Bfin = True
Unload Me
End Sub

Private Sub IAnimacion_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long

textoActual = ListRepetidos.List(ListRepetidos.listIndex)
bmpstring = ReadField(1, textoActual, Asc("."))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
        End If
        FormAuto.FrameAnim(0).Visible = True
        FormAuto.FrameAnim(1).Visible = False
        FormAuto.FrameAnim(2).Visible = False
        FormAuto.TextDatos(4).Text = BMPBuscado
    End If
End If
End Sub

Private Sub Ibody_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long

textoActual = ListRepetidos.List(ListRepetidos.listIndex)
bmpstring = ReadField(1, textoActual, Asc("."))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
        End If
        FormAuto.FrameAnim(1).Visible = True
        FormAuto.FrameAnim(0).Visible = False
        FormAuto.FrameAnim(2).Visible = False
        FormAuto.Combo2.Visible = False
        FormAuto.Labelbody.Visible = False
        FormAuto.Labelbody1.Visible = False
        FormAuto.Labelbody2.Visible = False
        FormAuto.Loff.Visible = False
        FormAuto.Loffx.Visible = False
        FormAuto.Loffy.Visible = False
        FormAuto.TextDatos2(7).Visible = False
        FormAuto.TextDatos2(8).Visible = False
        FormAuto.TextDatos2(0).Enabled = False
        FormAuto.TextDatos2(1).Enabled = False
        FormAuto.TextDatos2(6).Enabled = False
        FormAuto.Text1.Visible = False
        FormAuto.Text2.Visible = False
        FormAuto.CheckAuto.Visible = False
        FormAuto.Optiondimension(0).Visible = False
        FormAuto.Optiondimension(1).Visible = False
        FormAuto.Optiondimension(2).Visible = False
        FormAuto.Label5.Visible = False
        FormAuto.Label6.Visible = False
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(4).Text = BMPBuscado
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
        FormAuto.Optiondimension(0).value = True
        FormAuto.Labelbody.Visible = True
        FormAuto.Labelbody1.Visible = True
        FormAuto.Labelbody2.Visible = True
        FormAuto.Text1.Visible = True
        FormAuto.Text1.Enabled = False
        FormAuto.Text2.Visible = True
        FormAuto.Text2.Enabled = False
        FormAuto.CheckAuto.Visible = True
        FormAuto.CheckAuto.value = vbUnchecked
        FormAuto.Text1.Text = UBound(BodyData) + 1
        FormAuto.Text2.Text = "-38º0"
        FormAuto.Combo2.Visible = True
        FormAuto.Combo2.listIndex = 0
        FormAuto.Optiondimension(0).Visible = True
        FormAuto.Optiondimension(1).Visible = True
        FormAuto.Optiondimension(2).Visible = True
        FormAuto.Label5.Visible = True
        FormAuto.Label6.Visible = True
    End If
End If
 
End Sub

Private Sub ICompleto_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long

textoActual = ListRepetidos.List(ListRepetidos.listIndex)
bmpstring = ReadField(1, textoActual, Asc("."))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
        End If
        FormAuto.FrameAnim(0).Visible = False
        FormAuto.FrameAnim(1).Visible = False
        FormAuto.FrameAnim(2).Visible = True
        FormAuto.TextDatos3(4).Text = BMPBuscado
        FormAuto.TextDatos3(5).Text = BuscarGrHlibres(1)
    End If
End If
End Sub

Private Sub ListRepetidos_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long

textoActual = ListRepetidos.List(ListRepetidos.listIndex)
bmpstring = ReadField(1, textoActual, Asc("."))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
        Call frmMain.BuscarNuevoF(BMPBuscado)
    End If
End If

End Sub

Private Sub ListRepetidos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call Me.PopupMenu(Me.mnuopcionesI)
        
    End If
End Sub
