VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Abraxas • Map Editor"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   Begin VB.Timer backup 
      Interval        =   60000
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer tmrMinimap 
      Interval        =   500
      Left            =   720
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuAr 
      Caption         =   "Archivo"
      Begin VB.Menu mnuArchi 
         Caption         =   "Nuevo Mapa"
         Index           =   1
      End
      Begin VB.Menu mnuArchi 
         Caption         =   "Abrir mapa como ..."
         Index           =   2
      End
      Begin VB.Menu mnuArchi 
         Caption         =   "Guardar este mapa"
         Index           =   3
      End
      Begin VB.Menu mnuArchi 
         Caption         =   "Guardar mapa como ..."
         Index           =   4
      End
      Begin VB.Menu mnuArchi 
         Caption         =   "Salir"
         Index           =   5
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Edicion"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copiar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Pegar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "Deshacer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptCopy 
         Caption         =   "Opciones de Copiado"
         Begin VB.Menu mnuCSup 
            Caption         =   "Superficies"
            Begin VB.Menu mnuSupCap 
               Caption         =   "Capa 1"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuSupCap 
               Caption         =   "Capa 2"
               Checked         =   -1  'True
               Index           =   2
            End
            Begin VB.Menu mnuSupCap 
               Caption         =   "Capa 3"
               Checked         =   -1  'True
               Index           =   3
            End
            Begin VB.Menu mnuSupCap 
               Caption         =   "Capa 4"
               Checked         =   -1  'True
               Index           =   4
            End
         End
         Begin VB.Menu mnuCopObjs 
            Caption         =   "Objetos"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCopNpcs 
            Caption         =   "NPCs"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCopBloqs 
            Caption         =   "Bloqueos"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCopTrigs 
            Caption         =   "Triggers"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuAmd 
         Caption         =   "Administracion de Copiado"
      End
   End
   Begin VB.Menu mnuRender 
      Caption         =   "Renderizar Mapa"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public Change As Boolean
Public FileMap As String
Public FileMapDir As String
Public ViewX As Integer
Public ViewY As Integer


Private Sub backup_Timer()
    If (Dir$(App.Path & "\backups", vbDirectory) = vbNullString) Then
        MkDir App.Path & "\backups"
    End If
    
    Dim arch As String
    arch = App.Path & "\backups\[" & MapInfo.name & "] " & Day(Now) & "-" & Month(Now) & "  " & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".abr"
    
    Map_Save arch
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        bRunning = False
    ElseIf KeyCode = 46 Then
        If Shift = 1 Then
            
        End If
    ElseIf KeyCode = 110 Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        frmMain.Caption = MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyNumpad1 Then
        If PutSurface Then
            If frmMode.lstSuperfices.ListIndex > 0 Then frmMode.lstSuperfices.ListIndex = frmMode.lstSuperfices.ListIndex - 1
        End If
        
        If PutObjs Then
            If frmMode.ObjList.ListIndex > 0 Then frmMode.ObjList.ListIndex = frmMode.ObjList.ListIndex - 1
        End If
    ElseIf KeyCode = vbKeyNumpad0 Then
        If PutSurface Then
            If frmMode.lstSuperfices.ListIndex < frmMode.lstSuperfices.ListCount - 1 Then frmMode.lstSuperfices.ListIndex = frmMode.lstSuperfices.ListIndex + 1
        End If
        
        If PutObjs Then
            If frmMode.ObjList.ListIndex < frmMode.ObjList.ListCount - 1 Then frmMode.ObjList.ListIndex = frmMode.ObjList.ListIndex + 1
        End If
    End If
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    If Not InMapBounds(tX, tY) Then Exit Sub
    If Shift = 1 And Button = vbLeftButton Then
        If SelecAC = 0 Then
            SelecOX = tX
            SelecOY = tY
            If SelecOX <> 0 And SelecOY <> 0 Then
                InitGrh MapData(SelecOX, SelecOY).selec, 2
                SelecAC = 1
            End If
            
            If cData.dX > 0 Then Me.mnuPaste = True
        ElseIf SelecAC = 1 Then
            SelecDX = tX
            SelecDY = tY
                
            If tY < SelecOY Then
                SelecDY = SelecOY
                SelecOY = tY
            End If
            If tX < SelecOX Then
                SelecDX = SelecOX
                SelecOX = tX
            End If
            
            Dim x1 As Integer, y1 As Integer
            For x1 = SelecOX To SelecDX
                For y1 = SelecOY To SelecDY
                    If InMapBounds(x1, y1) Then
                        InitGrh MapData(x1, y1).selec, 2
                    End If
                Next y1
            Next x1
            SelecAC = 2
            frmMode.cmdSupInsertSelec.Enabled = True
            frmMode.cmdBlockInsertSelect.Enabled = True
            frmMode.cmdTrigInsertSelect.Enabled = True
            
            frmMode.cmdSupSacSelec.Enabled = True
            frmMode.cmdBlockSacSelect.Enabled = True
            frmMode.cmdTrigSacarSelect.Enabled = True
            frmMode.cmdSacarObjSelect.Enabled = True
            
            Me.mnuPaste = False
            Me.mnuCopy.Enabled = True
        ElseIf SelecAC = 2 Then
            For x1 = SelecOX To SelecDX
                For y1 = SelecOY To SelecDY
                    If InMapBounds(x1, y1) Then
                        MapData(x1, y1).selec.GrhIndex = 0
                    End If
                Next y1
            Next x1
            SelecAC = 0
            SelecOX = 0: SelecDX = 0: SelecOY = 0: SelecDY = 0
            
            frmMode.cmdSupInsertSelec.Enabled = False
            frmMode.cmdBlockInsertSelect.Enabled = False
            frmMode.cmdTrigInsertSelect.Enabled = False
            
            frmMode.cmdSupSacSelec.Enabled = False
            frmMode.cmdBlockSacSelect.Enabled = False
            frmMode.cmdTrigSacarSelect.Enabled = False
            frmMode.cmdSacarObjSelect.Enabled = False
            
            Me.mnuPaste = False
            Me.mnuCopy.Enabled = False
        End If
    ElseIf Shift = 1 And Button = vbRightButton Then
        If MapData(tX, tY).TileExit.map <> 0 Then
            frmMode.txtMap = MapData(tX, tY).TileExit.map
            frmMode.txtX = MapData(tX, tY).TileExit.x
            frmMode.txtY = MapData(tX, tY).TileExit.y
        End If
    End If
End Sub

Private Sub Form_Resize()
    Engine.Engine_Reset
End Sub

Private Sub mnuAmd_Click()
    frmAdmCopy.Show , Me
End Sub

Private Sub mnuArchi_Click(Index As Integer)
    Select Case Index
        Case 1
            If MsgBox("Desea guardar el mapa actual?", vbYesNo) = vbYes Then
                mnuArchi_Click 3
            End If
            
            ResetMap

        Case 2
            With CommonDialog1
                .Filter = "Archivos de mapas de Abraxas (*.abr) |*.abr|Archivos de mapas de Argentum Online (*.map) |*.map"
                
                .ShowOpen
                If .FileName <> vbNullString Then
                    If LCase$(Right$(.FileName, 4)) = ".abr" Then
                        Call Map_Load(.FileName)
                    ElseIf LCase$(Right$(.FileName, 4)) = ".map" Then
                        Call Map_Load_Old(Left$(.FileName, Len(.FileName) - 4))
                    End If
                End If
            End With
                
        Case 3
            If FileMap <> "" Then
                Call Map_Save(FileMap)
            Else
                mnuArchi_Click 4
            End If
            
        Case 4
            With CommonDialog1
                .Filter = "Archivos de mapas de Abraxas (*.abr) |*.abr"
                
                .ShowSave
                If .FileName <> vbNullString Then
                    If LCase$(Right$(.FileName, 4)) <> ".abr" Then
                        .FileName = .FileName & ".abr"
                    End If
                    
                    Call Map_Save(.FileName)
                End If
                
            End With
            
        Case 5
            Engine.Engine_Deinit
    End Select
End Sub

Private Sub mnuAuto_Click()

End Sub

Private Sub mnuCopy_Click()
    Dim x As Long, y As Long
    
    ReDim cData.copied(SelecDX - SelecOX + 1, SelecDY - SelecOY + 1)
    
    For x = 1 To SelecDX - SelecOX + 1
        For y = 1 To SelecDY - SelecOY + 1
            cData.copied(x, y) = MapData(SelecOX + x - 1, SelecOY + y - 1)
            
            cData.copied(x, y).light_value(0) = 0
            cData.copied(x, y).light_value(1) = 0
            cData.copied(x, y).light_value(2) = 0
            cData.copied(x, y).light_value(3) = 0
            
            cData.copied(x, y).selec.GrhIndex = 0
            
            cData.copied(x, y).Particle_Group = 0
            
            cData.copied(x, y).CharIndex = 0
        Next y
    Next x
    
    cData.dX = SelecDX - SelecOX + 1
    cData.dY = SelecDY - SelecOY + 1
    
    frmAdmCopy.cmdSaveCopy.Enabled = True
    
End Sub

Private Sub mnuPaste_Click()
    Dim x As Long, y As Long, i As Long
    
    modDeshacer.Deshacer_Add ""
    
    For x = 1 To cData.dX
        For y = 1 To cData.dY
            If cData.cNpc Then
                If cData.copied(x, y).NPCIndex <> 0 Then
                    Dim NPCIndex As Integer
                    NPCIndex = cData.copied(x, y).NPCIndex
    
                    Dim Body As Integer
                    Dim Head As Integer
                    Dim Heading As Byte
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call Engine.Char_Make(NextOpenChar(), Body, Head, Heading, SelecOX + x - 1, SelecOY + y - 1, 0, 0, 0)
                End If
            End If
            
            If cData.cBloq Then
                MapData(SelecOX + x - 1, SelecOY + y - 1).Blocked = cData.copied(x, y).Blocked
            End If
            
            If cData.cTrig Then
                MapData(SelecOX + x - 1, SelecOY + y - 1).Trigger = cData.copied(x, y).Trigger
            End If
            
            If cData.cObj Then
                MapData(SelecOX + x - 1, SelecOY + y - 1).ObjGrh = cData.copied(x, y).ObjGrh
                MapData(SelecOX + x - 1, SelecOY + y - 1).OBJInfo = cData.copied(x, y).OBJInfo
            End If
            
            For i = 1 To 4
                If cData.cCap(i) Then
                    MapData(SelecOX + x - 1, SelecOY + y - 1).Graphic(i) = cData.copied(x, y).Graphic(i)
                End If
            Next i
            
        Next y
    Next x
End Sub

Private Sub mnuDeshacer_Click()
    Call Deshacer_Recover
End Sub

Private Sub Form_Click()
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    UltimoClickX = tX
    UltimoClickY = tY
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

    If Shift = 1 Then Exit Sub
    
    If Not InMapBounds(LastPostCliked.x, LastPostCliked.y) Then Exit Sub
    
    If Button = vbRightButton Then
        If PutBlock Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Blocked = 0
        End If
         
        If PutSurface Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Graphic(frmMode.lstCapa.Text).GrhIndex = 0
            Change = True
        End If
         
        If PutAuto Then
            Deshacer_Add ""
            'MapData(LastPostCliked.X, LastPostCliked.Y).autoSelect = 0
            Change = True
        End If
         
        If PutObjs Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex = 0
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.amount = 0
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex = 0
            If frmMode.ckCap3.value = vbChecked Then
                MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3).GrhIndex = 0
            End If
            Change = True
        End If
         
        If PutParticles Then
            Deshacer_Add ""
            Engine.Particle_Delete MapData(LastPostCliked.x, LastPostCliked.y).Particle_Group
            MapData(LastPostCliked.x, LastPostCliked.y).Particle_index = 0
            Exit Sub
        End If
             
        If PutTrigger Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Trigger = 0
            Exit Sub
        End If
        
        If PutTrans Then
            Deshacer_Add ""
            MapData(tX, tY).TileExit.x = 0
            MapData(tX, tY).TileExit.y = 0
            MapData(tX, tY).TileExit.map = 0
        End If
        
        If PutNPC Then
            Deshacer_Add ""
            If GetKeyState(vbKeyControl) < 0 Then
                If Not MapData(tX, tY).NPCIndex = 0 Then frmMode.lstNPC.ListIndex = MapData(tX, tY).NPCIndex - 1
            Else
                Engine.Char_Erase MapData(tX, tY).CharIndex
                MapData(tX, tY).NPCIndex = 0
            End If
        End If
        
        If PutLight Then
            Deshacer_Add ""
            Engine.Light_Delete_From_Pos tX, tY
        End If
        
    ElseIf Button = vbLeftButton Then
        If PutAuto Then
            Deshacer_Add ""
            'MapData(LastPostCliked.X, LastPostCliked.Y).autoSelect = 1
            Change = True
        End If
        
        If PutLight Then
            Deshacer_Add ""
            Engine.Light_Create tX, tY, frmMode.txtRango.Text, D3DColorXRGB(frmMode.txtRed.Text, frmMode.txtGreen.Text, frmMode.txtBlue.Text)
        End If
        
        If PutNPC Then
            Deshacer_Add ""
            Dim NPCIndex As Integer
            NPCIndex = CInt(Val(ReadField(1, frmMode.lstNPC.Text, Asc("^"))))
            If NPCIndex = 0 Then Exit Sub
            Dim Body As Integer
            Dim Head As Integer
            Dim Heading As Byte
            Body = NpcData(NPCIndex).Body
            Head = NpcData(NPCIndex).Head
            Heading = NpcData(NPCIndex).Heading
            MapData(tX, tY).NPCIndex = NPCIndex
            Call Engine.Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, 0, 0, 0)
        End If

        If PutBlock Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Blocked = 1
            Exit Sub
        End If
         
        If PutSurface And Not frmMode.lstSuperfices.Text = "" Then
            Deshacer_Add ""
            Dim SurfaceIndex As Integer
            SurfaceIndex = CLng(ReadField(2, frmMode.lstSuperfices.Text, Asc("#")))
            
            If SurfaceIndex = 0 Then Exit Sub
            
            If SupData(SurfaceIndex).Width = 0 Then SupData(SurfaceIndex).Width = 1
            If SupData(SurfaceIndex).Height = 0 Then SupData(SurfaceIndex).Height = 1
                
            Dim aux As Integer
            Dim dY As Integer
            Dim dX As Integer
                         
            dY = 0
            dX = 0
                  
            If frmMode.AutoCompletarSuperficie.value = False Then
                Change = True
                aux = Val(SupData(SurfaceIndex).Grh) + _
                  (((tY + dY) Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) + ((tX + dX) Mod SupData(SurfaceIndex).Width)
                  
                If MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex <> aux Then
                    MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                    InitGrh MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)), aux
                End If
            Else
                Change = True
                
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, despTile As Integer
                
                tXX = tX
                tYY = tY
                
                If frmMode.chOrdenar = vbChecked Then
                    despTile = ((tY Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) _
                                + (tX Mod SupData(SurfaceIndex).Width)
                Else
                    despTile = 0
                End If
                
                For i = 1 To SupData(SurfaceIndex).Height
                    For j = 1 To SupData(SurfaceIndex).Width
                        aux = Val(SupData(SurfaceIndex).Grh) + despTile
                        MapData(tXX, tYY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                        InitGrh MapData(tXX, tYY).Graphic(Val(frmMode.lstCapa.Text)), aux
                        tXX = tXX + 1
                        despTile = despTile + 1
                        
                        If despTile = SupData(SurfaceIndex).Height * SupData(SurfaceIndex).Width Then despTile = 0
                    Next
                    tXX = tX
                    tYY = tYY + 1
                Next
                tYY = tY
            End If
        End If
        
        If PutObjs Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.amount = Val(frmMode.ObjCant.Text)
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex = frmMode.ObjClick
            
            If frmMode.ObjClick < 1 Then Exit Sub
            MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex = ObjData(MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex).GrhIndex
            Grh_Init MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh, MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex
 
            Change = True
        End If

        If frmMode.ckCap3.value = vbChecked And MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex <> 0 Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3).GrhIndex = ObjData(MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex).GrhIndex
            Grh_Init MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3), MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex
        End If
            
        If PutParticles Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Particle_index = frmMode.lstParticles.ListIndex + 1
            Engine.Particle_Save_Create frmMode.lstParticles.ListIndex + 1, MapData(LastPostCliked.x, LastPostCliked.y).Particle_Group
            Exit Sub
        End If
         
        If PutTrigger Then
            Deshacer_Add ""
            MapData(LastPostCliked.x, LastPostCliked.y).Trigger = frmMode.TrigList.ListIndex
            Exit Sub
        End If
        
        If PutTrans Then
            Deshacer_Add ""
            Dim map As Integer, XX As Integer, yy As Integer
            map = Val(frmMode.txtMap.Text): XX = Val(frmMode.txtX.Text): yy = Val(frmMode.txtY.Text)
            If XX > 99 Or XX < 1 Then
                MsgBox "Valor X invalido"
                Exit Sub
            End If
            
            If yy > 99 Or yy < 1 Then
                MsgBox "Valor Y invalido"
                Exit Sub
            End If
            
            If map > 500 Or map < 1 Then
                MsgBox "Valor Map invalido"
                Exit Sub
            End If
            
            MapData(tX, tY).TileExit.x = XX
            MapData(tX, tY).TileExit.y = yy
            MapData(tX, tY).TileExit.map = map
        End If
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    MiMouse = True
    Static tlX As Integer, tlY As Integer
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If Not InMapBounds(LastPostCliked.x, LastPostCliked.y) Then Exit Sub
    
    If Button = vbRightButton Then
         If PutBlock Then
             MapData(LastPostCliked.x, LastPostCliked.y).Blocked = 0
             Exit Sub
         End If
        
        If PutAuto Then
            'MapData(LastPostCliked.X, LastPostCliked.Y).autoSelect = 0
            Change = True
        End If
        
        If PutParticles Then
            Engine.Particle_Delete MapData(LastPostCliked.x, LastPostCliked.y).Particle_Group
            MapData(LastPostCliked.x, LastPostCliked.y).Particle_index = 0
            Exit Sub
        End If
         
        If PutSurface Then
            MapData(LastPostCliked.x, LastPostCliked.y).Graphic(frmMode.lstCapa.Text).GrhIndex = 0
            Change = True
        End If
        
         If PutObjs Then
            MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex = 0
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.amount = 0
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex = 0
            If frmMode.ckCap3.value = vbChecked Then
                MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3).GrhIndex = 0
            End If
            Change = True
         End If
         
         If PutTrigger Then
             MapData(LastPostCliked.x, LastPostCliked.y).Trigger = 0
             Exit Sub
         End If
         
        If PutTrans Then
            MapData(tX, tY).TileExit.x = 0
            MapData(tX, tY).TileExit.y = 0
            MapData(tX, tY).TileExit.map = 0
        End If
        
        If PutNPC Then
            If GetKeyState(vbKeyControl) < 0 Then
                If Not MapData(tX, tY).NPCIndex = 0 Then frmMode.lstNPC.ListIndex = MapData(tX, tY).NPCIndex - 1
            Else
                Engine.Char_Erase MapData(tX, tY).CharIndex
                MapData(tX, tY).NPCIndex = 0
            End If
        End If
        
        If PutLight Then
            Engine.Light_Delete_From_Pos tX, tY
        End If
        
    ElseIf Button = vbLeftButton Then
    
        If PutLight Then
            Engine.Light_Create tX, tY, frmMode.txtRango.Text, D3DColorXRGB(frmMode.txtRed.Text, frmMode.txtGreen.Text, frmMode.txtBlue.Text)
        End If
        
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
         If PutBlock Then
             MapData(LastPostCliked.x, LastPostCliked.y).Blocked = 1
             Exit Sub
         End If
        
        If PutAuto Then
            'MapData(LastPostCliked.X, LastPostCliked.Y).autoSelect = 1
            Change = True
        End If
        
        If PutNPC Then
            Dim NPCIndex As Integer
            NPCIndex = CInt(Val(ReadField(1, frmMode.lstNPC.Text, Asc("^"))))
            If NPCIndex = 0 Then Exit Sub
            Dim Body As Integer
            Dim Head As Integer
            Dim Heading As Byte
            Body = NpcData(NPCIndex).Body
            Head = NpcData(NPCIndex).Head
            Heading = NpcData(NPCIndex).Heading
            MapData(tX, tY).NPCIndex = NPCIndex
            Call Engine.Char_Make(NextOpenChar(), Body, Head, Heading, tX, tY, 0, 0, 0)
        End If
         
         If PutSurface And Not frmMode.lstSuperfices.Text = "" Then
            Dim SurfaceIndex As Integer
            SurfaceIndex = CLng(ReadField(2, frmMode.lstSuperfices.Text, Asc("#")))
            
            If SupData(SurfaceIndex).Width = 0 Then SupData(SurfaceIndex).Width = 1
            If SupData(SurfaceIndex).Height = 0 Then SupData(SurfaceIndex).Height = 1

            Dim aux As Integer
            Dim dY As Integer
            Dim dX As Integer
                         
            dY = 0
            dX = 0
                  
            If frmMode.AutoCompletarSuperficie.value = False Then
                Change = True
                aux = Val(SupData(SurfaceIndex).Grh) + _
                  (((tY + dY) Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) + ((tX + dX) Mod SupData(SurfaceIndex).Width)
                  
                If MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex <> aux Then
                    MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                    InitGrh MapData(tX, tY).Graphic(Val(frmMode.lstCapa.Text)), aux
                End If
            Else
                Change = True
                
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, despTile As Integer
                
                tXX = tX
                tYY = tY
                
                If frmMode.chOrdenar = vbChecked Then
                    despTile = ((tY Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) _
                                + (tX Mod SupData(SurfaceIndex).Width)
                Else
                    despTile = 0
                End If
                        
                For i = 1 To SupData(SurfaceIndex).Height
                    For j = 1 To SupData(SurfaceIndex).Width
                        aux = Val(SupData(SurfaceIndex).Grh) + despTile
                        If Not tXX > 98 Or Not tXX < 1 Or Not tYY > 98 Or Not tYY < 1 Then
                            If InMapBounds(tXX, tYY) Then
                                MapData(tXX, tYY).Graphic(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                                InitGrh MapData(tXX, tYY).Graphic(Val(frmMode.lstCapa.Text)), aux
                            End If
                            tXX = tXX + 1
                            despTile = despTile + 1
                            If despTile = SupData(SurfaceIndex).Height * SupData(SurfaceIndex).Width Then despTile = 0
                        
                        End If
                    Next
                    tXX = tX
                    tYY = tYY + 1
                Next
                tYY = tY
            End If
        End If
        
         If PutObjs Then
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.amount = Val(frmMode.ObjCant.Text)
            MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex = frmMode.ObjClick
            
            If frmMode.ObjClick = 0 Then Exit Sub
            MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex = ObjData(MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex).GrhIndex
             Grh_Init MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh, MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex
            
            Change = True
         End If
                     
        If frmMode.ckCap3.value = vbChecked And MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex <> 0 Then
            MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3).GrhIndex = ObjData(MapData(LastPostCliked.x, LastPostCliked.y).OBJInfo.ObjIndex).GrhIndex
            Grh_Init MapData(LastPostCliked.x, LastPostCliked.y).Graphic(3), MapData(LastPostCliked.x, LastPostCliked.y).ObjGrh.GrhIndex
        End If
            
         If PutParticles Then
             Engine.Particle_Save_Create frmMode.lstParticles.ListIndex + 1, MapData(LastPostCliked.x, LastPostCliked.y).Particle_Group
             Exit Sub
         End If
         
         If PutTrigger Then
             MapData(LastPostCliked.x, LastPostCliked.y).Trigger = frmMode.TrigList.ListIndex
             Exit Sub
         End If
         
        If PutTrans Then
            Dim map As Integer, XX As Integer, yy As Integer
            map = Val(frmMode.txtMap.Text): XX = Val(frmMode.txtX.Text): yy = Val(frmMode.txtY.Text)
            If XX > 99 Or XX < 1 Then
                MsgBox "Valor X invalido"
                Exit Sub
            End If
            
            If yy > 99 Or yy < 1 Then
                MsgBox "Valor Y invalido"
                Exit Sub
            End If
            
            If map > 500 Or map < 1 Then
                MsgBox "Valor Map invalido"
                Exit Sub
            End If
            
            MapData(tX, tY).TileExit.x = XX
            MapData(tX, tY).TileExit.y = yy
            MapData(tX, tY).TileExit.map = map
        End If
    Else
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        Dim xx1 As Integer, yy1 As Integer
        If PutObjs Then
            If (tlX <> tX And tlX <> 0) Or (tlY <> tY And tlY <> 0) Then
                MapData(tlX, tlY).objView.GrhIndex = 0
            End If
            
            If frmMode.ObjClick = -1 Or frmMode.ObjClick = 0 Then Exit Sub
            MapData(tX, tY).objView.GrhIndex = ObjData(frmMode.ObjClick).GrhIndex
            
            Grh_Init MapData(tX, tY).objView, MapData(tX, tY).objView.GrhIndex
            tlX = tX
            tlY = tY
        End If
        
        If PutSurface And Not frmMode.lstSuperfices.Text = "" Then
            If tX <> ViewX Or tY <> ViewY Then
                If ViewY <> 0 Or ViewX <> 0 Then
                    For xx1 = ViewX To ViewX + 8
                        For yy1 = ViewY To ViewY + 8
                            If InMapBounds(xx1, yy1) Then MapData(xx1, yy1).gView(Val(frmMode.lstCapa.Text)).GrhIndex = 0
                        Next yy1
                    Next xx1
                End If
            Else
                Exit Sub
            End If
  
            If frmMode.lstSuperfices.Text = "" Then Exit Sub
            SurfaceIndex = CLng(ReadField(2, frmMode.lstSuperfices.Text, Asc("#")))
            
            If SupData(SurfaceIndex).Width = 0 Then SupData(SurfaceIndex).Width = 1
            If SupData(SurfaceIndex).Height = 0 Then SupData(SurfaceIndex).Height = 1

            dY = 0
            dX = 0
                  
            If frmMode.AutoCompletarSuperficie.value = False Then
                Change = True
                aux = Val(SupData(SurfaceIndex).Grh) + _
                  (((tY + dY) Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) + ((tX + dX) Mod SupData(SurfaceIndex).Width)
                  
                If InMapBounds(tX, tY) Then
                    If MapData(tX, tY).gView(Val(frmMode.lstCapa.Text)).GrhIndex <> aux Then
                        MapData(tX, tY).gView(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                        InitGrh MapData(tX, tY).gView(Val(frmMode.lstCapa.Text)), aux
                    End If
                End If
            Else
                Change = True
                
                'Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                
                If frmMode.chOrdenar = vbChecked Then
                    despTile = ((tY Mod SupData(SurfaceIndex).Height) * SupData(SurfaceIndex).Width) _
                                + (tX Mod SupData(SurfaceIndex).Width)
                Else
                    despTile = 0
                End If
                
                tXX = tX
                tYY = tY
                
                For i = 1 To SupData(SurfaceIndex).Height
                    For j = 1 To SupData(SurfaceIndex).Width
                        aux = Val(SupData(SurfaceIndex).Grh) + despTile
                        
                        If InMapBounds(tXX, tYY) Then
                            MapData(tXX, tYY).gView(Val(frmMode.lstCapa.Text)).GrhIndex = aux
                            InitGrh MapData(tXX, tYY).gView(Val(frmMode.lstCapa.Text)), aux
                        End If
                        tXX = tXX + 1
                        despTile = despTile + 1
                        
                        If despTile = SupData(SurfaceIndex).Height * SupData(SurfaceIndex).Width Then despTile = 0
                    Next
                    tXX = tX
                    tYY = tYY + 1
                Next
                tYY = tY
                
            End If
        End If
        ViewX = tX
        ViewY = tY
    End If
End Sub

Function ResetMap()
    Engine.Light_Delete_All
    Engine.Particle_Delete 1, 1
    
    Dim i As Long
    For i = 1 To 10000
        Engine.Char_Erase i
    Next i
    
    FileMap = ""
    
    Map_New
    
End Function

Private Sub mnuRender_Click()
    frmRender.Show , Me
    Unload frmRender
End Sub

Private Sub mnuSupCap_Click(Index As Integer)
    mnuSupCap(Index).Checked = Not mnuSupCap(Index).Checked
    cData.cCap(Index) = Not cData.cCap(Index)
End Sub

Private Sub mnuCopBloqs_Click()
    mnuCopBloqs.Checked = Not mnuCopBloqs.Checked
    cData.cBloq = Not cData.cBloq
End Sub

Private Sub mnuCopNpcs_Click()
    mnuCopNpcs.Checked = Not mnuCopNpcs.Checked
    cData.cNpc = Not cData.cNpc
End Sub

Private Sub mnuCopObjs_Click()
    mnuCopObjs.Checked = Not mnuCopObjs.Checked
    cData.cObj = Not cData.cObj
End Sub

Private Sub mnuCopTrigs_Click()
    mnuCopTrigs.Checked = Not mnuCopTrigs.Checked
    cData.cTrig = Not cData.cTrig
End Sub
