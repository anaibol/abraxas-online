VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost

Dim PrevTempSlot As Byte

Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Public Sub SelectGold()
'Sets the gold as the currently selected Item
    Dim prevSelectedItem As Long
    
    'Store preivously selected Item
    prevSelectedItem = InvSelSlot
    
    'Select the gold
    InvSelSlot = FLAGORO
    
    'Redraw old Item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawSlot(prevSelectedItem)
    End If
End Sub

Public Sub DeselectItem()
'Deselect the currently selected Item
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelSlot
    
    'Select nothing
    InvSelSlot = 0
    
    'Redraw old Item to deselect it
    Call DrawSlot(ClearSlot)

End Sub

Public Property Get SelectedItem() As Long
'Retrieves the selected Item index
    
    Select Case InventoryWindow.Tag
            
        Case frmMain.PicInv.Tag
            SelectedItem = InvSelSlot
            
        Case frmMain.PicSpellInv.Tag
            SelectedItem = SpellSelSlot
            
        Case frmMain.PicBelt.Tag
            SelectedItem = BeltSelSlot
        
        'Case frmMain.PicCompaInv.Tag
        '    SelectedItem = CompaSelSlot
        
        Case frmComerciar.PicComercianteInv.Tag, frmBanco.PicBancoInv.Tag
            SelectedItem = NpcInvSelSlot
            
    End Select
End Property

Public Function ClickItem(ByVal x As Integer, ByVal y As Integer) As Long
'Selects the Item clicked if it's valid and return's it's index

    Dim Item As Byte
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    Item = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth / 32) + 1
    
    'Make sure it's within limits
    If Item <= UBound(Inv) Then
        ClickItem = Item
    End If
End Function

Public Sub DrawSlot(ByVal Slot As Integer)
'Renders a inventory Slot to the given PictureBox
        
On Error GoTo Error

    If Slot < 1 Then
        Exit Sub
    End If

    Dim GrhIndex As Integer
    Dim Amount As Integer
    Dim PuedeUsar As Boolean
    
    Select Case InventoryWindow.Tag
        
        Case frmMain.PicInv.Tag
            With Inv(Slot)
                GrhIndex = .GrhIndex
                Amount = .Amount
                PuedeUsar = .PuedeUsar
            End With
            
        Case frmComerciar.PicComercianteInv.Tag
            With NpcInv(Slot)
                GrhIndex = .GrhIndex
                Amount = .Amount
                PuedeUsar = .PuedeUsar
            End With
            
        Case frmBanco.PicBancoInv.Tag
            With Banco(Slot)
                GrhIndex = .GrhIndex
                Amount = .Amount
                PuedeUsar = .PuedeUsar
            End With

    End Select
    
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    
    With SrcRect
        .Right = 32
        .Bottom = 32
    End With
        
    With InvSurface
        .BltFast 0, 0, SurfaceDB.Surface(2), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
    
        If GrhIndex > 0 Then
            'Render the Item grh
            .BltFast 0, 0, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            
            If AlphaBActivated Then
                If Not PuedeUsar Then
                    Call EfectoInventario(0, 0, Slot, GrhIndex)
                End If
                
                If (InventoryWindow.Tag = frmMain.PicInv.Tag And Slot = InvSelSlot) Or _
                (InventoryWindow.Tag = frmComerciar.PicComercianteInv.Tag And Slot = NpcInvSelSlot) Or _
                (InventoryWindow.Tag = frmBanco.PicBancoInv.Tag And Slot = NpcInvSelSlot) Then
    
                    If PuedeUsar Then
                        Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                    End If
                    
                    .SetForeColor RGB(200, 150, 70)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    
                ElseIf (InventoryWindow.Tag = frmMain.PicInv.Tag And Slot = TempSlot) Or _
                (InventoryWindow.Tag = frmComerciar.PicComercianteInv.Tag And Slot = NpcTempSlot) Or _
                (InventoryWindow.Tag = frmBanco.PicBancoInv.Tag And Slot = NpcTempSlot) Then
    
                    If PuedeUsar Then
                        Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
                    End If
                    
                    .SetForeColor RGB(100, 70, 50)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
        
                If Amount > 1 Then
                
                    Call .SetFont(frmCantidad.font)
                    
                    .SetForeColor vbBlack
                    .DrawText -1, 0, Amount, False
                    .DrawText 1, 0, Amount, False
                    .DrawText 0, 1, Amount, False
                    .DrawText 0, -1, Amount, False
        
                    If PuedeUsar Then
                        .SetForeColor &HC0FFFF
                        .DrawText 0, 0, PonerPuntos(Amount), False
                    Else
                        .SetForeColor vbRed
                        .DrawText 0, 0, PonerPuntos(Amount), False
                    End If
                End If
            Else
                If (InventoryWindow.Tag = frmMain.PicInv.Tag And Slot = InvSelSlot) Or _
                (InventoryWindow.Tag = frmComerciar.PicComercianteInv.Tag And Slot = NpcInvSelSlot) Or _
                (InventoryWindow.Tag = frmBanco.PicBancoInv.Tag And Slot = NpcInvSelSlot) Then
                    
                    .SetForeColor RGB(200, 150, 0)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
                ElseIf Slot = TempSlot Then
                    .SetForeColor RGB(255, 190, 70)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
                
                'NO PUEDE USAR
                If Not PuedeUsar Then
                    .SetForeColor RGB(200, 0, 0)
                    .DrawText 20, 20, "x", False
                End If
                
                If Amount > 1 Then
                    Call .SetFont(frmCantidad.font)
                    .SetForeColor vbBlack
                    .DrawText -1, 0, Amount, False
                    .DrawText 1, 0, Amount, False
                    .DrawText 0, 1, Amount, False
                    .DrawText 0, -1, Amount, False
        
                    If PuedeUsar Then
                        .SetForeColor &HC0FFFF
                        .DrawText 0, 0, PonerPuntos(Amount), False
                    Else
                        .SetForeColor vbRed
                        .DrawText 0, 0, PonerPuntos(Amount), False
                    End If
                End If
            End If
        End If
        
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
        
    End With
   
    Call InventoryWindow.Refresh
    
Error:
End Sub

Public Sub DrawBeltSlot(ByVal Slot As Integer)

On Error GoTo Error

    If Slot < 1 Then
        Exit Sub
    End If

    Dim GrhIndex As Integer
    Dim Amount As Integer
    
    GrhIndex = Belt(Slot).GrhIndex
    Amount = Belt(Slot).Amount
    
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With

    With SrcRect
        .Right = 32
        .Bottom = 32
    End With
    
    With InvSurface
        .BltFast 0, 0, SurfaceDB.Surface(2), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
    
        If GrhIndex > 0 Then
            'Render the Item grh
            .BltFast 0, 0, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            
            If AlphaBActivated Then
            
                If Not Belt(Slot).PuedeUsar Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 0, 0, 0)
                    
                    If Slot = BeltTempSlot Then
                        .SetForeColor RGB(255, 190, 70)
                        
                        .setDrawStyle DrawStyleConstants.vbSolid
                        .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    End If
                
                ElseIf Slot = BeltSelSlot Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                    
                    .SetForeColor RGB(200, 150, 70)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    
                ElseIf Slot = BeltTempSlot Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
    
                    .SetForeColor RGB(100, 70, 50)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
            
            Else
                If Not Belt(Slot).PuedeUsar Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 0, 0, 0)
                    
                    'If Slot = BeltTempSlot Then
                        .SetForeColor RGB(90, 90, 50)
                    
                        .setDrawStyle DrawStyleConstants.vbSolid
                        .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    'End If
            
                ElseIf Slot = BeltSelSlot Then
                    .SetForeColor RGB(200, 150, 0)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
                ElseIf Slot = BeltTempSlot Then
                    .SetForeColor RGB(255, 190, 70)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
    
            End If
            
            If Amount > 1 Then
            
                Call .SetFont(frmCantidad.font)
                
                .SetForeColor vbBlack
                .DrawText -1, 0, Amount, False
                .DrawText 1, 0, Amount, False
                .DrawText 0, 1, Amount, False
                .DrawText 0, -1, Amount, False
    
                If Belt(Slot).PuedeUsar Then
                    .SetForeColor &HC0FFFF
                    .DrawText 0, 0, Amount, False
                Else
                    .SetForeColor vbRed
                    .DrawText 0, 0, Amount, False
                End If
            End If
            
            Call .SetFont(frmCantidad.font)
            
            .SetForeColor vbBlack
            .DrawText 25, 21, Slot, False
            .DrawText 27, 21, Slot, False
            .DrawText 26, 22, Slot, False
            .DrawText 26, 20, Slot, False
    
            If Belt(Slot).PuedeUsar Then
                .SetForeColor &H808080
                .DrawText 26, 21, Slot, False
            Else
                .SetForeColor &H80&
                .DrawText 26, 21, Slot, False
            End If
        End If
        
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
        
    End With
   
    Call InventoryWindow.Refresh
    
Error:
End Sub

Public Sub DrawSpellSlot(ByVal Slot As Integer)

On Error GoTo Error
    
    If Slot < 1 Then
        Exit Sub
    End If
    
    Dim GrhIndex As Integer
    
    GrhIndex = Spell(Slot).Grh
    
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    
    With SrcRect
        .Right = 32
        .Bottom = 32
    End With
    
    With InvSurface
        If GrhIndex < 1 Then
            .BltFast 0, 0, SurfaceDB.Surface(2), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        
        Else
            'Render the Item grh
            .BltFast 0, 0, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
            
            If AlphaBActivated Then
            
                If Not Spell(Slot).PuedeLanzar Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 0, 0, 0)
                    
                    'If Slot = TempSlot Then
                        .SetForeColor RGB(90, 90, 50)
                    
                        .setDrawStyle DrawStyleConstants.vbSolid
                        .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    'End If
                
                ElseIf Slot = SpellSelSlot Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                    
                    .SetForeColor RGB(200, 150, 70)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    
                ElseIf Slot = TempSlot Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
                        
                    .SetForeColor RGB(90, 90, 50)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
            
            Else
                If Not Spell(Slot).PuedeLanzar Then
                
                    Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 0, 0, 0)
                    
                    'If Slot = TempSlot Then
                        .SetForeColor RGB(90, 90, 50)
                    
                        .setDrawStyle DrawStyleConstants.vbSolid
                        .DrawRoundedBox 0, 0, 32, 32, 5, 5
                    'End If
            
                ElseIf Slot = SpellSelSlot Then
                    .SetForeColor RGB(200, 150, 0)
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
                ElseIf Slot = TempSlot Then
                    .SetForeColor RGB(255, 190, 70)
                    
                    .setDrawStyle DrawStyleConstants.vbSolid
                    .DrawRoundedBox 0, 0, 32, 32, 5, 5
                End If
            
            End If
         End If
         
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
        
    End With
   
    Call InventoryWindow.Refresh
    
Error:
End Sub

Public Sub DrawCompaSlotBody(ByVal Slot As Integer, ByVal GrhIndex As Integer)
'Renders a inventory Slot to the given PictureBox
    
On Error GoTo Error
    
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    
    With InvSurface
        'Error ACA - IMAGEN FONDO Slot
        '.BltFast 0, 0, SurfaceDB.Surface(123456789), destRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            
        'TONCES PONGO ESTO
        'Clear the Slot area
        Call .BltColorFill(TempRect, InventoryWindow.BackColor)
    
        'BODY
        'Get source rect
        With GrhData(GrhData(BodyData(Compa(Slot).Body).Walk(3).GrhIndex).Frames(1))
            SrcRect.Left = .sX
            SrcRect.Top = .sY
            SrcRect.Right = SrcRect.Left + 32 - 10
            SrcRect.Bottom = SrcRect.Top + 32 - 9
        End With
                      
        'Render the Item grh
        .BltFast 3, 9, SurfaceDB.Surface(GrhData(GrhData(BodyData(Compa(Slot).Body).Walk(3).GrhIndex).Frames(1)).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        
        If AlphaBActivated Then
            If Slot = CompaSelSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                
                .SetForeColor RGB(200, 150, 70)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
            ElseIf Slot = CompaTempSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
                    
                .SetForeColor RGB(90, 90, 50)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If
        Else
            If Slot = CompaSelSlot Then
                .SetForeColor RGB(200, 150, 0)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            ElseIf Slot = CompaTempSlot Then
                .SetForeColor RGB(255, 190, 70)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If

        End If
        
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
    End With
   
    'Call InventoryWindow.Refresh
    
    Call DrawCompaSlotHead(Slot, GrhIndex)
    
Error:
End Sub

Public Sub DrawCompaSlotHead(ByVal Slot As Integer, ByVal GrhIndex As Integer)
'Renders a inventory Slot to the given PictureBox

On Error GoTo Error

    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    
    With InvSurface
        'Error ACA - IMAGEN FONDO Slot
        '.BltFast 0, 0, SurfaceDB.Surface(123456789), destRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            
        'TONCES PONGO ESTO
        'Clear the Slot area
        'Call .BltColorFill(TempRect, InventoryWindow.BackColor)
                    
        'HEAD
        'Get source rect
        With GrhData(HeadData(Compa(Slot).Head).Head(3).GrhIndex)
            SrcRect.Left = .sX
            SrcRect.Top = .sY
            SrcRect.Right = SrcRect.Left + 32 - 16
            SrcRect.Bottom = SrcRect.Top + 32 + BodyData(Compa(Slot).Body).HeadOffset.y
        End With
        
        'Render the Item grh
        .BltFast 7, 3, SurfaceDB.Surface(GrhData(GrhData(HeadData(Compa(Slot).Head).Head(3).GrhIndex).Frames(1)).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

        If AlphaBActivated Then
            If Slot = CompaSelSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                
                .SetForeColor RGB(200, 150, 70)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
            ElseIf Slot = CompaTempSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
                    
                .SetForeColor RGB(90, 90, 50)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If
        Else
            If Slot = CompaSelSlot Then
                .SetForeColor RGB(200, 150, 0)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            ElseIf Slot = CompaTempSlot Then
                .SetForeColor RGB(255, 190, 70)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If

        End If
        
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
    End With
   
    Call DrawCompaSlotCasco(Slot, GrhIndex)

Error:
End Sub

Public Sub DrawCompaSlotCasco(ByVal Slot As Integer, ByVal GrhIndex As Integer)
'Renders a inventory Slot to the given PictureBox
    
    If Compa(Slot).CascoAnim < 3 Then
        Exit Sub
    End If
    
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim DestRect As RECT
   
    With TempRect
        .Bottom = 32
        .Right = 32
    End With
   
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
   
    'If not in renderable area we exit
    If Slot < upperLeftSlot Then
        Exit Sub
    End If
   
    With DestRect
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    
    With InvSurface
        'Error ACA - IMAGEN FONDO Slot
        '.BltFast 0, 0, SurfaceDB.Surface(123456789), destRect, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
            
        'TONCES PONGO ESTO
        'Clear the Slot area
        'Call .BltColorFill(TempRect, InventoryWindow.BackColor)
                    
        'HEAD
        'Get source rect
        With GrhData(CascoAnimData(Compa(Slot).CascoAnim).Head(3).GrhIndex)
            SrcRect.Left = .sX
            SrcRect.Top = .sY
            SrcRect.Right = SrcRect.Left + 32 - 16
            SrcRect.Bottom = SrcRect.Top + 32 + BodyData(Compa(Slot).Body).HeadOffset.y
        End With
        
        'Render the Item grh
        .BltFast 7, 3, SurfaceDB.Surface(GrhData(GrhData(CascoAnimData(Compa(Slot).CascoAnim).Head(3).GrhIndex).Frames(1)).FileNum), SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

        If AlphaBActivated Then
            If Slot = CompaSelSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 200, 200, 60)
                
                .SetForeColor RGB(200, 150, 70)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
                
            ElseIf Slot = CompaTempSlot Then
            
                Call DrawTonalityEffect(Slot, GrhIndex, 0, 0, 150, 150, 40)
                    
                .SetForeColor RGB(90, 90, 50)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If
        Else
            If Slot = CompaSelSlot Then
                .SetForeColor RGB(200, 150, 0)
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            ElseIf Slot = CompaTempSlot Then
                .SetForeColor RGB(255, 190, 70)
                
                .setDrawStyle DrawStyleConstants.vbSolid
                .DrawRoundedBox 0, 0, 32, 32, 5, 5
            End If

        End If
        
        'Render the Item to the Inventory Window
        .BltToDC InventoryWindow.hdc, TempRect, DestRect
    End With
   
    Call InventoryWindow.Refresh
End Sub

Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                        Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                        Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                        Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                        Optional ByVal bShowText As Boolean = True)

    Set InventoryWindow = InvPic
    
    'Make sure auto-Redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
        
    'initialize DX stuff
    Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectDraw was correctly initialized
    If DirectDraw Is Nothing Then
        Exit Sub
    End If
    
    'Set surface's description
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = 32
        .lWidth = 32
    End With

    'Create inventory surface
    Set InvSurface = DirectDraw.CreateSurface(SurfaceDesc)
    
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error Resume Next

    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then
        Exit Sub
    End If

    If PicInvDragging Then
        Exit Sub
    End If
    
    If frmCarp.Visible Then
        Exit Sub
    End If
    
    If frmHerrero.Visible Then
        Exit Sub
    End If
                
    Select Case InventoryWindow.Tag
    
        Case frmMain.PicInv.Tag
            
            If TempSlot > 0 Then
                If Inv(TempSlot).ObjIndex > 0 Then
                    PrevTempSlot = TempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            TempSlot = ClickItem(CInt(x), CInt(y))
    
            If TempSlot < 1 Or TempSlot > MaxInvSlots Then
                Exit Sub
            End If
            
            If TempSlot <> PrevTempSlot Then
                If PrevTempSlot > 0 Then
                    Call DrawSlot(PrevTempSlot)
                End If
            End If
            
            If Inv(TempSlot).ObjIndex < 1 Then
                TempSlot = 0

                frmMain.Label.Caption = vbNullString
                InventoryWindow.ToolTipText = vbNullString
                                
                If PrevTempSlot > 0 Then
                    Call DrawSlot(PrevTempSlot)
                End If
                
                Exit Sub
            End If
            
            Call DrawSlot(TempSlot)
            
            If Not Comerciando Then
                frmMain.Label.Caption = Inv(TempSlot).Name
                
                Select Case Inv(TempSlot).ObjType
                    Case 2, 32
                        InventoryWindow.ToolTipText = "Daño: " & Inv(TempSlot).MinHit & " - " & Inv(TempSlot).MaxHit
                    Case 3, 16, 17
                        InventoryWindow.ToolTipText = "Defensa: " & Inv(TempSlot).MinDef & " - " & Inv(TempSlot).MaxDef
                    Case Else
                        InventoryWindow.ToolTipText = vbNullString
                End Select
                
            ElseIf frmComerciar.Visible Then
            
                With frmComerciar
                    Exit Sub
                    If Inv(TempSlot).ObjIndex > 0 Then
                        .lblItemName.Caption = Inv(TempSlot).Name
                        .lblItemPrice.Caption = PonerPuntos(CalculateSellPrice(Inv(TempSlot).Valor, Val(Replace(frmComerciar.Cantidad.Text, ".", vbNullString))))
                        .lblItemName.Visible = True
                        .lblItemPrice.Visible = True
                        .lblValor.Visible = True
    
                        Select Case Inv(TempSlot).ObjType
                            Case 2, 32
                                .lblDamage.Caption = "Daño:"
                                .lblMinHit.Caption = Inv(TempSlot).MinHit
                                .lblMaxHit.Caption = Inv(TempSlot).MaxHit
                                
                                .lblDamage.Visible = True
                                .lblMinHit.Visible = True
                                .lblMaxHit.Visible = True
                                
                                .lblDefense.Visible = False
                                .lblMinDef.Visible = False
                                .lblMaxDef.Visible = False
                                
                                .lblGuion.Visible = True
                                .lblGuion2.Visible = False
                            Case 3, 16, 17
                                .lblDefense.Caption = "Defensa:"
                                .lblMinDef.Caption = Inv(TempSlot).MinDef
                                .lblMaxDef.Caption = Inv(TempSlot).MaxDef
                                
                                .lblDefense.Visible = True
                                .lblMinDef.Visible = True
                                .lblMaxDef.Visible = True
                                
                                .lblDamage.Visible = False
                                .lblMinHit.Visible = False
                                .lblMaxHit.Visible = False
                                
                                .lblGuion.Visible = False
                                .lblGuion2.Visible = True
                            Case Else
                                .lblDamage.Visible = False
                                .lblMinHit.Visible = False
                                .lblMaxHit.Visible = False
                                
                                .lblDefense.Visible = False
                                .lblMinDef.Visible = False
                                .lblMaxDef.Visible = False
                                
                                .lblGuion.Visible = False
                                .lblGuion2.Visible = False
                        End Select
                        
                        If Inv(TempSlot).PuedeUsar = False Then
                            .lblPuedeUsar.Visible = True
                        Else
                            .lblPuedeUsar.Visible = False
                        End If
                        
                        .lblNpcName.Visible = False
                    Else
                        .lblItemName.Visible = False
                        .lblValor.Visible = False
                        .lblItemPrice.Visible = False
                        .lblPuedeUsar.Visible = False
                        
                        .lblDamage.Visible = False
                        .lblMinHit.Visible = False
                        .lblMaxHit.Visible = False
                                                
                        .lblDefense.Visible = False
                        .lblMinDef.Visible = False
                        .lblMaxDef.Visible = False
                        
                        .lblGuion.Visible = False
                        .lblGuion2.Visible = False
                        
                        .lblNpcName.Visible = True
                    End If
                End With
                    
            ElseIf frmBanco.Visible Then
        
                With frmBanco
                    
                    If Inv(TempSlot).ObjIndex > 0 Then
                        .lblItemName.Caption = Inv(TempSlot).Name
                        .lblItemPrice.Caption = PonerPuntos(CalculateSellPrice(Inv(TempSlot).Valor, Val(Replace(frmComerciar.Cantidad.Text, ".", vbNullString))))
                        .lblItemName.Visible = True
                        .lblItemPrice.Visible = True
                        .lblValor.Visible = True
    
                        Select Case Inv(TempSlot).ObjType
                            Case 2, 32
                                .lblDamage.Caption = "Daño:"
                                .lblMinHit.Caption = Inv(TempSlot).MinHit
                                .lblMaxHit.Caption = Inv(TempSlot).MaxHit
                                
                                .lblDamage.Visible = True
                                .lblMinHit.Visible = True
                                .lblMaxHit.Visible = True
                                
                                .lblDefense.Visible = False
                                .lblMinDef.Visible = False
                                .lblMaxDef.Visible = False
                                
                                .lblGuion.Visible = True
                                .lblGuion2.Visible = False
                            Case 3, 16, 17
                                .lblDefense.Caption = "Defensa:"
                                .lblMinDef.Caption = Inv(TempSlot).MinDef
                                .lblMaxDef.Caption = Inv(TempSlot).MaxDef
                                
                                .lblDefense.Visible = True
                                .lblMinDef.Visible = True
                                .lblMaxDef.Visible = True
                                
                                .lblDamage.Visible = False
                                .lblMinHit.Visible = False
                                .lblMaxHit.Visible = False
                                
                                .lblGuion.Visible = False
                                .lblGuion2.Visible = True
                            Case Else
                                .lblDamage.Visible = False
                                .lblMinHit.Visible = False
                                .lblMaxHit.Visible = False
                                
                                .lblDefense.Visible = False
                                .lblMinDef.Visible = False
                                .lblMaxDef.Visible = False
                                
                                .lblGuion.Visible = False
                                .lblGuion2.Visible = False
                        End Select
                        
                        If Inv(TempSlot).PuedeUsar = False Then
                            .lblPuedeUsar.Visible = True
                        Else
                            .lblPuedeUsar.Visible = False
                        End If
                        
                        .lblBoveda.Visible = False
                    Else
                        .lblItemName.Visible = False
                        .lblValor.Visible = False
                        .lblItemPrice.Visible = False
                        .lblPuedeUsar.Visible = False
                        
                        .lblDamage.Visible = False
                        .lblMinHit.Visible = False
                        .lblMaxHit.Visible = False
                                                
                        .lblDefense.Visible = False
                        .lblMinDef.Visible = False
                        .lblMaxDef.Visible = False
                        
                        .lblGuion.Visible = False
                        .lblGuion2.Visible = False
                        
                        .lblBoveda.Visible = True
                    End If
                End With
                
            End If
        
        Case frmMain.PicBelt.Tag
            
            If BeltTempSlot > 0 Then
                If Belt(BeltTempSlot).ObjIndex > 0 Then
                    PrevTempSlot = BeltTempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            BeltTempSlot = ClickItem(CInt(x), CInt(y))
    
            If BeltTempSlot < 1 Or BeltTempSlot > MaxBeltSlots Then
                Exit Sub
            End If
                        
            If BeltTempSlot <> PrevTempSlot Then
                If PrevTempSlot > 0 Then
                    Call DrawBeltSlot(PrevTempSlot)
                End If
            End If
                 
            If Belt(BeltTempSlot).ObjIndex < 1 Then
                BeltTempSlot = 0
                frmMain.Label.Caption = vbNullString
                InventoryWindow.ToolTipText = vbNullString
                                
                If PrevTempSlot > 0 Then
                    Call DrawBeltSlot(PrevTempSlot)
                End If
            End If

            Call DrawBeltSlot(BeltTempSlot)

            If Not Comerciando Then
                frmMain.Label.Caption = Belt(BeltTempSlot).Name
                
                'Select Case Belt(BeltTempSlot).ObjType
                '    Case 2, 32
                '        InventoryWindow.ToolTipText = "Daño: " & Belt(BeltTempSlot).MinHit & " - " & Belt(BeltTempSlot).MaxHit
                '    Case 3, 16, 17
                '        InventoryWindow.ToolTipText = "Defensa: " & Belt(BeltTempSlot).MinDef & " - " & Belt(BeltTempSlot).MaxDef
                '    Case Else
                '        InventoryWindow.ToolTipText = vbNullString
                'End Select
                
            ElseIf frmComerciar.Visible Then
            
                With frmComerciar

                    If Belt(BeltTempSlot).ObjIndex > 0 Then
                        .lblItemName.Caption = Belt(BeltTempSlot).Name
                        .lblItemPrice.Caption = PonerPuntos(CalculateSellPrice(Belt(BeltTempSlot).Valor, Val(Replace(frmComerciar.Cantidad.Text, ".", vbNullString))))
                        .lblItemName.Visible = True
                        .lblItemPrice.Visible = True
                        .lblValor.Visible = True
    
                        'Select Case Belt(BeltTempSlot).ObjType
                        '    Case 2, 32
                        '        .lblDamage.Caption = "Daño:"
                        '        .lblMinHit.Caption = Belt(BeltTempSlot).MinHit
                        '        .lblMaxHit.Caption = Belt(BeltTempSlot).MaxHit
                                
                        '        .lblDamage.Visible = True
                        '        .lblMinHit.Visible = True
                        '        .lblMaxHit.Visible = True
                                
                        '        .lblDefense.Visible = False
                        '        .lblMinDef.Visible = False
                        '        .lblMaxDef.Visible = False
                                
                        '        .lblGuion.Visible = True
                        '        .lblGuion2.Visible = False
                        '    Case 3, 16, 17
                        '        .lblDefense.Caption = "Defensa:"
                        '        .lblMinDef.Caption = Belt(BeltTempSlot).MinDef
                        '        .lblMaxDef.Caption = Belt(BeltTempSlot).MaxDef
                                
                        '        .lblDefense.Visible = True
                        '        .lblMinDef.Visible = True
                        '        .lblMaxDef.Visible = True
                                
                        '        .lblDamage.Visible = False
                        '        .lblMinHit.Visible = False
                        '        .lblMaxHit.Visible = False
                                
                        '        .lblGuion.Visible = False
                        '        .lblGuion2.Visible = True
                        '    Case Else
                        '        .lblDamage.Visible = False
                        '        .lblMinHit.Visible = False
                        '        .lblMaxHit.Visible = False
                                
                        '        .lblDefense.Visible = False
                        '        .lblMinDef.Visible = False
                        '        .lblMaxDef.Visible = False
                                
                        '        .lblGuion.Visible = False
                        '        .lblGuion2.Visible = False
                        'End Select
                                                    
                        .lblPuedeUsar.Visible = False

                        .lblNpcName.Visible = False
                    Else
                        .lblItemName.Visible = False
                        .lblValor.Visible = False
                        .lblItemPrice.Visible = False
                        .lblPuedeUsar.Visible = False
                        
                        .lblDamage.Visible = False
                        .lblMinHit.Visible = False
                        .lblMaxHit.Visible = False
                                                
                        .lblDefense.Visible = False
                        .lblMinDef.Visible = False
                        .lblMaxDef.Visible = False
                        
                        .lblGuion.Visible = False
                        .lblGuion2.Visible = False
                        
                        .lblNpcName.Visible = True
                    End If
                End With
                    
            ElseIf frmBanco.Visible Then
        
                With frmBanco
                    
                    If Belt(BeltTempSlot).ObjIndex > 0 Then
                        .lblItemName.Caption = Belt(BeltTempSlot).Name
                        .lblItemPrice.Caption = PonerPuntos(CalculateSellPrice(Belt(BeltTempSlot).Valor, Val(Replace(frmComerciar.Cantidad.Text, ".", vbNullString))))
                        .lblItemName.Visible = True
                        .lblItemPrice.Visible = True
                        .lblValor.Visible = True
    
                        'Select Case Belt(BeltTempSlot).ObjType
                        '    Case 2, 32
                        '       .lblDamage.Caption = "Daño:"
                        '       .lblMinHit.Caption = Belt(BeltTempSlot).MinHit
                        '       .lblMaxHit.Caption = Belt(BeltTempSlot).MaxHit
                                
                        '       .lblDamage.Visible = True
                        '       .lblMinHit.Visible = True
                        '       .lblMaxHit.Visible = True
                                
                        '       .lblDefense.Visible = False
                        '       .lblMinDef.Visible = False
                        '       .lblMaxDef.Visible = False
                                
                        '       .lblGuion.Visible = True
                        '       .lblGuion2.Visible = False
                         '   Case 3, 16, 17
                        '       .lblDefense.Caption = "Defensa:"
                        '       .lblMinDef.Caption = Belt(BeltTempSlot).MinDef
                        '       .lblMaxDef.Caption = Belt(BeltTempSlot).MaxDef
                                
                        '       .lblDefense.Visible = True
                        '       .lblMinDef.Visible = True
                        '       .lblMaxDef.Visible = True
                                
                        '       .lblDamage.Visible = False
                        '       .lblMinHit.Visible = False
                        '       .lblMaxHit.Visible = False
                                
                        '       .lblGuion.Visible = False
                        '       .lblGuion2.Visible = True
                         '   Case Else
                        '       .lblDamage.Visible = False
                        '       .lblMinHit.Visible = False
                        '       .lblMaxHit.Visible = False
                                
                        '       .lblDefense.Visible = False
                        '       .lblMinDef.Visible = False
                        '       .lblMaxDef.Visible = False
                                
                        '       .lblGuion.Visible = False
                        '       .lblGuion2.Visible = False
                        'End Select
                        
                        .lblPuedeUsar.Visible = False
    
                        .lblBoveda.Visible = False
                    Else
                        .lblItemName.Visible = False
                        .lblValor.Visible = False
                        .lblItemPrice.Visible = False
                        .lblPuedeUsar.Visible = False
                        
                        .lblDamage.Visible = False
                        .lblMinHit.Visible = False
                        .lblMaxHit.Visible = False
                                                
                        .lblDefense.Visible = False
                        .lblMinDef.Visible = False
                        .lblMaxDef.Visible = False
                        
                        .lblGuion.Visible = False
                        .lblGuion2.Visible = False
                        
                        .lblBoveda.Visible = True
                    End If
                End With
                
            End If
        
        Case frmMain.PicSpellInv.Tag
        
            If TempSlot > 0 Then
                If Spell(TempSlot).Grh > 0 Then
                    PrevTempSlot = TempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            TempSlot = ClickItem(CInt(x), CInt(y))
        
            If TempSlot < 1 Or TempSlot > MaxSpellSlots Then
                Exit Sub
            End If
            
            If TempSlot <> PrevTempSlot Then
                If PrevTempSlot > 0 Then
                    Call DrawSpellSlot(PrevTempSlot)
                End If
            End If
            
            If Spell(TempSlot).Grh > 0 Then
                Call DrawSpellSlot(TempSlot)
                frmMain.Label.Caption = Spell(TempSlot).Nombre
            Else
                frmMain.Label.Caption = vbNullString
                TempSlot = 0
            End If

            
        Case frmMain.PicCompaInv.Tag
            
            If CompaTempSlot > 0 Then
                If Compa(CompaTempSlot).Body > 0 Then
                    PrevTempSlot = CompaTempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            CompaTempSlot = ClickItem(CInt(x), CInt(y))
            
            If CompaTempSlot < 1 Or CompaTempSlot > MaxCompaSlots Then
                Exit Sub
            End If
            
           If CompaTempSlot <> PrevTempSlot Then
                If PrevTempSlot > 0 Then
                    Call DrawCompaSlotBody(PrevTempSlot, GrhData(BodyData(Compa(PrevTempSlot).Body).Walk(3).GrhIndex).Frames(1))
                End If
            End If
            
            Call DrawCompaSlotBody(CompaTempSlot, GrhData(BodyData(Compa(CompaTempSlot).Body).Walk(3).GrhIndex).Frames(1))
            
            If Compa(CompaTempSlot).Body < 1 Then
                CompaTempSlot = 0
                Exit Sub
            End If
                        
            If Button = vbRightButton Then
                If CompaSelSlot > 0 Then
                    If CompaTempSlot <> CompaSelSlot Then
                        If PicInvDragging Then
                            If InventoryWindow.Tag = frmMain.PicCompaInv.Tag Then
                                If Compa(CompaSelSlot).Body > 0 Then
                                    PicInvDragging = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        Case frmComerciar.PicComercianteInv.Tag
            
            If NpcTempSlot > 0 Then
                If NpcInv(NpcTempSlot).ObjIndex > 0 Then
                    PrevTempSlot = NpcTempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            NpcTempSlot = ClickItem(CInt(x), CInt(y))
        
            If NpcTempSlot < 1 Or NpcTempSlot > MaxNpcInvSlots Then
                Exit Sub
            End If
            
            If NpcTempSlot <> PrevTempSlot Then
                If PrevTempSlot > 0 Then
                    Call DrawSlot(PrevTempSlot)
                End If
            End If
            
            If NpcInv(NpcTempSlot).ObjIndex < 1 Then
                NpcTempSlot = 0
                Exit Sub
            End If
            
            Call DrawSlot(NpcTempSlot)
            
            With frmComerciar
                If NpcInv(NpcTempSlot).ObjIndex > 0 Then
                    .lblItemName.Caption = NpcInv(NpcTempSlot).Name
                    .lblItemPrice.Caption = PonerPuntos(CalculateBuyPrice(NpcInv(NpcTempSlot).Valor, Val(frmComerciar.Cantidad.Text)))
                    .lblItemName.Visible = True
                    .lblItemPrice.Visible = True
                    .lblValor.Visible = True

                    Select Case NpcInv(NpcTempSlot).ObjType
                    
                        Case 2, 32
                            .lblDamage.Caption = "Daño:"
                            .lblMinHit.Caption = NpcInv(NpcTempSlot).MinHit
                            .lblMaxHit.Caption = NpcInv(NpcTempSlot).MaxHit
                            
                            .lblDamage.Visible = True
                            .lblMinHit.Visible = True
                            .lblMaxHit.Visible = True
                            
                            .lblDefense.Visible = False
                            .lblMinDef.Visible = False
                            .lblMaxDef.Visible = False
                            
                            .lblGuion.Visible = True
                            .lblGuion2.Visible = False
                            
                        Case 3, 16, 17
                            .lblDefense.Caption = "Defensa:"
                            .lblMinDef.Caption = NpcInv(NpcTempSlot).MinDef
                            .lblMaxDef.Caption = NpcInv(NpcTempSlot).MaxDef
                            
                            .lblDefense.Visible = True
                            .lblMinDef.Visible = True
                            .lblMaxDef.Visible = True
                            
                            .lblDamage.Visible = False
                            .lblMinHit.Visible = False
                            .lblMaxHit.Visible = False
                            
                            .lblGuion.Visible = False
                            .lblGuion2.Visible = True
                            
                        Case Else
                            .lblDamage.Visible = False
                            .lblMinHit.Visible = False
                            .lblMaxHit.Visible = False
                            
                            .lblDefense.Visible = False
                            .lblMinDef.Visible = False
                            .lblMaxDef.Visible = False
                            
                            .lblGuion.Visible = False
                            .lblGuion2.Visible = False
                    End Select
                    
                    If NpcInv(NpcTempSlot).PuedeUsar = False Then
                        .lblPuedeUsar.Visible = True
                    Else
                        .lblPuedeUsar.Visible = False
                    End If
                    
                    .lblNpcName.Visible = False
                    
                Else
                    .lblItemName.Visible = False
                    .lblValor.Visible = False
                    .lblItemPrice.Visible = False
                    .lblPuedeUsar.Visible = False
                    
                    .lblDamage.Visible = False
                    .lblMinHit.Visible = False
                    .lblMaxHit.Visible = False
                                            
                    .lblDefense.Visible = False
                    .lblMinDef.Visible = False
                    .lblMaxDef.Visible = False
                    
                    .lblGuion.Visible = False
                    .lblGuion2.Visible = False
                    
                    .lblNpcName.Visible = True
                End If
            End With
        
        Case frmBanco.PicBancoInv.Tag
        
            If NpcTempSlot > 0 Then
                If NpcInv(NpcTempSlot).ObjIndex > 0 Then
                    PrevTempSlot = NpcTempSlot
                Else
                    PrevTempSlot = 0
                End If
            End If
            
            NpcTempSlot = ClickItem(CInt(x), CInt(y))
        
            If NpcTempSlot < 1 Or NpcTempSlot > MaxNpcInvSlots Then
                Exit Sub
            End If
            
            If Button = vbRightButton Then
                If NpcTempSlot <> NpcInvSelSlot Then
                    'If 'PicInvMouseDown Then
                    '    If NpcInv(NpcInvSelSlot).ObjIndex > 0 Then
                    '        PicInvDragging = True
                            'PicInvMouseDown = False
                    '    End If
                    'End If
                End If
            End If
        
            If NpcInv(NpcTempSlot).ObjIndex > 0 Then
                Call DrawSlot(NpcTempSlot)
            End If
            
            If PrevTempSlot > 0 Then
                Call DrawSlot(PrevTempSlot)
            End If
        
            With frmBanco
                If NpcInv(NpcTempSlot).ObjIndex > 0 Then
                    .lblItemName.Caption = NpcInv(NpcTempSlot).Name
                    .lblItemPrice.Caption = PonerPuntos(CalculateBuyPrice(NpcInv(NpcTempSlot).Valor, Val(frmBanco.Cantidad.Text)))
                    .lblItemName.Visible = True
                    .lblItemPrice.Visible = True
                    .lblValor.Visible = True

                    Select Case NpcInv(NpcTempSlot).ObjType
                        Case 2, 32
                            .lblDamage.Caption = "Daño:"
                            .lblMinHit.Caption = NpcInv(NpcTempSlot).MinHit
                            .lblMaxHit.Caption = NpcInv(NpcTempSlot).MaxHit
                            
                            .lblDamage.Visible = True
                            .lblMinHit.Visible = True
                            .lblMaxHit.Visible = True
                            
                            .lblDefense.Visible = False
                            .lblMinDef.Visible = False
                            .lblMaxDef.Visible = False
                            
                            .lblGuion.Visible = True
                            .lblGuion2.Visible = False
                        Case 3, 16, 17
                            .lblDefense.Caption = "Defensa:"
                            .lblMinDef.Caption = NpcInv(NpcTempSlot).MinDef
                            .lblMaxDef.Caption = NpcInv(NpcTempSlot).MaxDef
                            
                            .lblDefense.Visible = True
                            .lblMinDef.Visible = True
                            .lblMaxDef.Visible = True
                            
                            .lblDamage.Visible = False
                            .lblMinHit.Visible = False
                            .lblMaxHit.Visible = False
                            
                            .lblGuion.Visible = False
                            .lblGuion2.Visible = True
                        Case Else
                            .lblDamage.Visible = False
                            .lblMinHit.Visible = False
                            .lblMaxHit.Visible = False
                            
                            .lblDefense.Visible = False
                            .lblMinDef.Visible = False
                            .lblMaxDef.Visible = False
                            
                            .lblGuion.Visible = False
                            .lblGuion2.Visible = False
                    End Select
                    
                    If NpcInv(NpcTempSlot).PuedeUsar = False Then
                        .lblPuedeUsar.Visible = True
                    Else
                        .lblPuedeUsar.Visible = False
                    End If
                    
                    .lblBoveda.Visible = False
                Else
                    .lblItemName.Visible = False
                    .lblValor.Visible = False
                    .lblItemPrice.Visible = False
                    .lblPuedeUsar.Visible = False
                    
                    .lblDamage.Visible = False
                    .lblMinHit.Visible = False
                    .lblMaxHit.Visible = False
                                            
                    .lblDefense.Visible = False
                    .lblMinDef.Visible = False
                    .lblMaxDef.Visible = False
                    
                    .lblGuion.Visible = False
                    .lblGuion2.Visible = False
                    
                    .lblBoveda.Visible = True
                End If
            End With
    End Select
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If x > 0 And y > 0 And x < InventoryWindow.Width And y < InventoryWindow.Height Then
            
        Dim PrevSlot As Byte
        Dim posList As Integer
        
        Select Case InventoryWindow.Tag
        
            Case frmMain.PicInv.Tag
                                        
                If InvSelSlot > 0 Then
                    PrevSlot = InvSelSlot
                End If
                
                InvSelSlot = ClickItem(CInt(x), CInt(y))
                
                If InvSelSlot < 1 Or InvSelSlot > MaxInvSlots Then
                    Exit Sub
                End If
        
                If PrevSlot > 0 Then
                    If InvSelSlot <> PrevSlot Then
                        Call DrawSlot(PrevSlot)
                    End If
                End If
                
                If Inv(InvSelSlot).ObjIndex < 1 Then
                    InvSelSlot = 0
                    Exit Sub
                End If

                Call DrawSlot(InvSelSlot)
                
                If Not Comerciando Then
                    If Button = vbRightButton Then
                        Dim i As Integer
                        
                        If Inv(InvSelSlot).GrhIndex < 1 Then
                            Exit Sub
                        End If
                        
                        PicInvDragging = True
    
                        last_i = InvSelSlot
                        
                        If last_i > 0 And last_i <= MaxInvSlots Then
                            posList = browseName(Inv(InvSelSlot).GrhIndex)
                            
                            If posList = 0 Then
                                i = GrhData(Inv(InvSelSlot).GrhIndex).FileNum
                
                                 DoEvents
                                 i = GrhData(Inv(InvSelSlot).GrhIndex).FileNum
                                 frmMain.MouseImage.ListImages.Add , CStr("g" & Inv(InvSelSlot).GrhIndex), Picture:=LoadPicture(App.path & "\Grh\" & (GrhData(Inv(InvSelSlot).GrhIndex).FileNum & ".bmp"))
                                 posList = frmMain.MouseImage.ListImages.Count
                            End If
                            
                            Set frmMain.PicInv.MouseIcon = frmMain.MouseImage.ListImages(posList).ExtractIcon
                            frmMain.PicInv.MousePointer = vbCustom
                            
                            Exit Sub
                        End If
                    End If
                End If

            Case frmMain.PicBelt.Tag
            
                If BeltSelSlot > 0 Then
                    PrevSlot = BeltSelSlot
                End If
            
                BeltSelSlot = ClickItem(CInt(x), CInt(y))
            
                If BeltSelSlot < 1 Or BeltSelSlot > MaxBeltSlots Then
                    Exit Sub
                End If
            
                If PrevSlot > 0 Then
                    If BeltSelSlot <> PrevSlot Then
                        Call DrawBeltSlot(PrevSlot)
                    End If
                End If
                
                If Belt(BeltSelSlot).ObjIndex < 1 Then
                    BeltSelSlot = 0
                    Exit Sub
                End If
                
                If Not Belt(BeltSelSlot).PuedeUsar Then
                    BeltSelSlot = 0
                    Exit Sub
                End If
                
                Call DrawBeltSlot(BeltSelSlot)
            
                 If Button = vbRightButton Then
                    PicInvDragging = True
                    
                    last_i = BeltSelSlot
                    
                    If last_i > 0 And last_i <= MaxBeltSlots Then
                        posList = browseName(Belt(BeltSelSlot).GrhIndex)
                        
                        If posList = 0 Then
                            i = GrhData(Belt(BeltSelSlot).GrhIndex).FileNum
            
                             DoEvents
                             i = GrhData(Belt(BeltSelSlot).GrhIndex).FileNum
                             frmMain.MouseImage.ListImages.Add , CStr("g" & Belt(BeltSelSlot).GrhIndex), Picture:=LoadPicture(GrhPath & (GrhData(Belt(BeltSelSlot).GrhIndex).FileNum & ".bmp"))
                             posList = frmMain.MouseImage.ListImages.Count
                        End If
                        
                        Set frmMain.PicBelt.MouseIcon = frmMain.MouseImage.ListImages(posList).ExtractIcon
                        frmMain.PicBelt.MousePointer = vbCustom
            
                        Exit Sub
                    End If
                End If
        
            Case frmMain.PicSpellInv.Tag
            
                If SpellSelSlot > 0 Then
                    PrevSlot = SpellSelSlot
                End If
                
                SpellSelSlot = ClickItem(CInt(x), CInt(y))
                    
                If SpellSelSlot < 1 Or SpellSelSlot > MaxSpellSlots Then
                    Exit Sub
                End If
                
                If PrevSlot > 0 Then
                    If SpellSelSlot <> PrevSlot Then
                        Call DrawSpellSlot(PrevSlot)
                    End If
                End If
                                    
                If Spell(SpellSelSlot).Grh < 1 Then
                    SpellSelSlot = 0
                    Exit Sub
                End If
                
                If Not Spell(SpellSelSlot).PuedeLanzar Then
                    SpellSelSlot = 0
                    Exit Sub
                End If
                
                Call DrawSpellSlot(SpellSelSlot)
                
                If Button = vbRightButton Then
                    PicInvDragging = True
                    
                    last_i = SpellSelSlot
                    
                    If last_i > 0 And last_i <= MaxSpellSlots Then
                        posList = browseName(Spell(SpellSelSlot).Grh)
                        
                        If posList = 0 Then
                            i = GrhData(Spell(SpellSelSlot).Grh).FileNum
            
                             DoEvents
                             i = GrhData(Spell(SpellSelSlot).Grh).FileNum
                             frmMain.MouseImage.ListImages.Add , CStr("g" & Spell(SpellSelSlot).Grh), Picture:=LoadPicture(GrhPath & (GrhData(Spell(SpellSelSlot).Grh).FileNum & ".bmp"))
                             posList = frmMain.MouseImage.ListImages.Count
                        End If
                        
                        Set frmMain.PicSpellInv.MouseIcon = frmMain.MouseImage.ListImages(posList).ExtractIcon
                        frmMain.PicSpellInv.MousePointer = vbCustom
            
                        Exit Sub
                    End If
                End If
                
            Case frmMain.PicCompaInv.Tag
                
                If CompaSelSlot > 0 Then
                    PrevSlot = CompaSelSlot
                End If
                
                CompaSelSlot = ClickItem(CInt(x), CInt(y))
                
                If CompaSelSlot < 1 Or CompaSelSlot > MaxCompaSlots Then
                    Exit Sub
                End If
                
                If Compa(CompaSelSlot).Body < 1 Then
                    CompaSelSlot = 0
                End If
            
                If PrevSlot > 0 Then
                    If PrevSlot <> CompaSelSlot Then
                        Call DrawCompaSlotBody(PrevSlot, GrhData(BodyData(Compa(PrevSlot).Body).Walk(3).GrhIndex).Frames(1))
                    End If
                End If
                
            Case frmComerciar.PicComercianteInv.Tag
            
                If NpcInvSelSlot > 0 Then
                    If NpcInv(NpcInvSelSlot).ObjIndex > 0 Then
                        PrevSlot = NpcInvSelSlot
                    Else
                        PrevSlot = 0
                    End If
                End If
                
                NpcInvSelSlot = ClickItem(CInt(x), CInt(y))
                
                If NpcInvSelSlot < 1 Or NpcInvSelSlot > MaxNpcInvSlots Then
                    Exit Sub
                End If
            
                If PrevSlot <> NpcInvSelSlot Then
                    If NpcInv(NpcInvSelSlot).ObjIndex > 0 Then
                        Call DrawSlot(PrevSlot)
                    Else
                        NpcInvSelSlot = 0
                    End If
                    
                    If PrevSlot > 0 Then
                        Call DrawSlot(PrevSlot)
                    End If
                End If
                
            Case frmBanco.PicBancoInv.Tag
                
                If NpcInvSelSlot > 0 Then
                    If Banco(NpcInvSelSlot).ObjIndex > 0 Then
                        PrevSlot = NpcInvSelSlot
                    Else
                        PrevSlot = 0
                    End If
                End If
                
                NpcInvSelSlot = ClickItem(CInt(x), CInt(y))
                
                If NpcInvSelSlot < 1 Or NpcInvSelSlot > MaxNpcInvSlots Then
                    Exit Sub
                End If
            
                If PrevSlot <> NpcInvSelSlot Then
                    If Banco(NpcInvSelSlot).ObjIndex > 0 Then
                        Call DrawSlot(PrevSlot)
                    Else
                        NpcInvSelSlot = 0
                    End If
                    
                    If PrevSlot > 0 Then
                        Call DrawSlot(PrevSlot)
                    End If
                End If
                
        End Select

    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then
        If PicInvDragging Then
            PicInvDragging = False
            frmMain.PicInv.MousePointer = vbDefault
        End If
        Exit Sub
    End If
    
    Dim TempSlot As Byte
    
    Select Case InventoryWindow.Tag
    
        Case frmMain.PicInv.Tag
        
           If PicInvDragging Then
               If Button = vbRightButton Then
                   TempSlot = ClickItem(CInt(x), CInt(y))
        
                   If InvSelSlot <> TempSlot Then
                       If InvSelSlot > 0 And TempSlot > 0 And TempSlot <= MaxInvSlots Then
                           Call MoveInvSlot(InvSelSlot, TempSlot)
                           Call Audio.Play(SND_CLICK)
                       End If
                   End If
               End If
               
               PicInvDragging = False
               frmMain.PicInv.MousePointer = vbDefault
           End If
           
        Case frmMain.PicBelt.Tag
        
           If PicInvDragging Then
               If Button = vbRightButton Then
                   TempSlot = ClickItem(CInt(x), CInt(y))
        
                   If BeltSelSlot <> TempSlot Then
                       If BeltSelSlot > 0 And TempSlot > 0 And TempSlot <= MaxBeltSlots Then
                           Call MoveBeltSlot(BeltSelSlot, TempSlot)
                           Call Audio.Play(SND_CLICK)
                       End If
                   End If
               End If
               
               PicInvDragging = False
               frmMain.PicBelt.MousePointer = vbDefault
           End If
           
        Case frmMain.PicSpellInv.Tag
        
           If PicInvDragging Then
               If Button = vbRightButton Then
                   TempSlot = ClickItem(CInt(x), CInt(y))
        
                   If SpellSelSlot <> TempSlot Then
                       If SpellSelSlot > 0 And TempSlot > 0 And TempSlot <= MaxSpellSlots Then
                            If Spell(SpellSelSlot).Grh > 0 Then
                                Call MoveSpellSlotSlot(SpellSelSlot, TempSlot)
                                Call Audio.Play(SND_CLICK)
                            End If
                       End If
                   End If
               End If
               
               PicInvDragging = False
               frmMain.PicSpellInv.MousePointer = vbDefault
           End If
           
        Case frmMain.PicCompaInv.Tag
        
            Exit Sub
            
            If PicInvDragging Then
                If Button = vbRightButton Then
                    TempSlot = ClickItem(CInt(x), CInt(y))
            
                    If CompaSelSlot <> TempSlot Then
                        If CompaSelSlot > 0 And TempSlot > 0 And TempSlot <= MaxCompaSlots Then
                            'Call MoveSlot(CompaSelSlot, TempSlot)
                            Call Audio.Play(SND_CLICK)
                        End If
                    End If
                End If
                
                PicInvDragging = False
                frmMain.PicCompaInv.MousePointer = vbDefault
            End If
    
        Case frmComerciar.PicComercianteInv.Tag
            Call Auto_Drag(frmComerciar.hWnd)
            
        Case frmBanco.PicBancoInv.Tag
            Call Auto_Drag(frmBanco.hWnd)
            
    End Select
    
    'ElseIf DragType = None Then
    'If Button = vbRightButton Then
    'If frmComerciar.Visible Then
    'WriteCommerceEnd
    'ElseIf frmBanco.Visible Then
    'WriteBankEnd
    'End If
    'End If
    'End If
    'Else
    'If InventoryWindow.Tag = frmMain.PicInv.Tag Then
    'If DragType = MiInventario Then
    'End If
    'ElseIf InventoryWindow.Tag = frmBanco.PicBancoInv.Tag Or frmComerciar.PicComercianteInv.Tag Then
    'DragType = InventarioNpc
    'End If
    'End If
End Sub

Private Sub MoveInvSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    
    If Slot2 > 0 Then
    
        Dim ObjIndex As Integer
        Dim Name As String
        Dim Amount As Integer
        Dim GrhIndex As Integer
        Dim ObjType As eObjType
        Dim MaxHit As Integer
        Dim MinHit As Integer
        Dim MaxDef As Byte
        Dim MinDef As Byte
        Dim Valor As Long
        Dim PuedeUsar As Boolean
        Dim Proyectil As Boolean

        With Inv(Slot2)
            ObjIndex = .ObjIndex
            Name = .Name
            Amount = .Amount
            GrhIndex = .GrhIndex
            ObjType = .ObjType
            MaxHit = .MaxHit
            MinHit = .MinHit
            MaxDef = .MaxDef
            MinDef = .MinDef
            Valor = .Valor
            PuedeUsar = .PuedeUsar
            Proyectil = .Proyectil
        End With

        With Inv(Slot)
            Call Inventario.SetSlot(Slot2, .ObjIndex, .Amount, .GrhIndex, .ObjType, .MinHit, .MaxHit, .MinDef, .MaxDef, .Valor, .Name, .PuedeUsar, .Proyectil)
        End With
        
        Call Inventario.SetSlot(Slot, ObjIndex, Amount, GrhIndex, ObjType, MinHit, MaxHit, MinDef, MaxDef, Valor, Name, PuedeUsar, Proyectil)
                
        Call WriteMoveInvSlot(Slot, Slot2)
        InvSelSlot = Slot2

    Else
        With Inv(Slot)
            Call Inventario.SetSlot(Slot2, .ObjIndex, .Amount, .GrhIndex, .ObjType, .MinHit, .MaxHit, .MinDef, .MaxDef, .Valor, .Name, .PuedeUsar, .Proyectil)
        End With
        
        Call Inventario.UnSetSlot(Slot)
    
        Call WriteMoveInvSlot(Slot, 0)
        InvSelSlot = Slot2
    End If

End Sub

Private Sub MoveBeltSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    
    If Slot2 > 0 Then
    
        Dim ObjIndex As Integer
        Dim Name As String
        Dim Amount As Integer
        Dim GrhIndex As Integer
        Dim Valor As Long
        Dim PuedeUsar As Boolean

        With Belt(Slot2)
            ObjIndex = .ObjIndex
            Name = .Name
            Amount = .Amount
            GrhIndex = .GrhIndex
            Valor = .Valor
            PuedeUsar = .PuedeUsar
        End With

        With Belt(Slot)
            Call SetBeltSlot(Slot2, .ObjIndex, .Amount, .GrhIndex, .Valor, .Name, .PuedeUsar)
        End With
        
        Call SetBeltSlot(Slot, ObjIndex, Amount, GrhIndex, Valor, Name, PuedeUsar)
                
        Call WriteMoveBeltSlot(Slot, Slot2)
        BeltSelSlot = Slot2

    Else
        With Belt(Slot)
            Call SetBeltSlot(Slot2, .ObjIndex, .Amount, .GrhIndex, .Valor, .Name, .PuedeUsar)
        End With
        
        Call Cinturon.UnSetSlot(Slot)
    
        Call WriteMoveBeltSlot(Slot, 0)
        BeltSelSlot = Slot2
    End If

End Sub

Private Sub MoveSpellSlotSlot(ByVal Slot As Byte, ByVal Slot2 As Byte)
    
    If Spell(Slot2).Grh > 0 Then
        
        Dim Grh As Integer
        Dim Nombre As String
        Dim MinSkill As Boolean
        Dim ManaRequerido As Integer
        Dim StaRequerido As Integer
        Dim NeedStaff As Integer
        Dim PuedeLanzar As Boolean
        
        With Spell(Slot2)
            Grh = .Grh
            Nombre = .Nombre
            MinSkill = .MinSkill
            ManaRequerido = .ManaRequerido
            StaRequerido = .StaRequerido
            NeedStaff = .NeedStaff
            PuedeLanzar = .PuedeLanzar
        End With
        
        With Spell(Slot)
            Call SetSpellSlot(Slot2, .Grh, .Nombre, .MinSkill, .ManaRequerido, .StaRequerido, .NeedStaff)
        End With
        
        Call SetSpellSlot(Slot, Grh, Nombre, MinSkill, ManaRequerido, StaRequerido, NeedStaff)
        
        Call WriteMoveSpellSlot(Slot, Slot2)
    
        SpellSelSlot = 0
        
        Call DrawSpellSlot(Slot)
    Else
    
        With Spell(Slot)
            Call SetSpellSlot(Slot2, .Grh, .Nombre, .MinSkill, .ManaRequerido, .StaRequerido, .NeedStaff)
        End With
        
        Call Hechizos.UnSetSlot(Slot)
        
        Call WriteMoveSpellSlot(Slot, Slot2)
        
        SpellSelSlot = 0
        
        Call DrawSpellSlot(Slot)
    End If

End Sub

Private Sub EfectoInventario(ByVal x As Integer, ByVal y As Integer, ByVal Slot As Integer, ByVal GrhIndex As Integer)
'Draw the GRH with a tonality, in the position X and Y of the Inventory

    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
 
    With SourceRect
        .Left = GrhData(GrhIndex).sX
        .Top = GrhData(GrhIndex).sY
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With

    Dim upperLeftSlot As Integer
    
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
 
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
   
    Set Src = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
   
    Src.GetSurfaceDesc ddsdSrc
    InvSurface.GetSurfaceDesc ddsdDest
   
    With rDest
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
   
    SrcLock = False
    DstLock = False
   
On Error GoTo HayErrorAlpha
   
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call InvSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
   
    Call InvSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
 
'Alpha
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 100, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 150, 0, 0)
    
    InvSurface.Unlock rDest
    Src.Unlock SourceRect
Exit Sub
 
HayErrorAlpha:
    If SrcLock Then
        Src.Unlock SourceRect
    End If
    
    If DstLock Then
        InvSurface.Unlock rDest
    End If
End Sub

Private Sub EfectoInventario2(ByVal x As Integer, ByVal y As Integer, ByVal Slot As Integer, ByVal GrhIndex As Integer)
    
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
 
        With SourceRect
            .Left = GrhData(GrhIndex).sX
            .Top = GrhData(GrhIndex).sY
            .Right = .Left + 32
            .Bottom = .Top + 32
        End With
 
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
 
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
   
    Set Src = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
   
    Src.GetSurfaceDesc ddsdSrc
    InvSurface.GetSurfaceDesc ddsdDest
   
    With rDest
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
      
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    Call InvSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
   
    Call InvSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
 
    Call vbDABLcolorblend16565(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 50, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0, 255, 0)
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 50, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0, 0, 0)
    
    InvSurface.Unlock rDest
    Src.Unlock SourceRect
End Sub

Private Sub DrawTonalityEffect(ByVal Slot As Integer, ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)
'Draw the GRH with a tonality, in the position X and Y of the Inventory
    
    Dim SourceRect As RECT
 
    If GrhIndex < 1 Then
        Exit Sub
    End If
    
    With SourceRect
        .Left = GrhData(GrhIndex).sX
        .Top = GrhData(GrhIndex).sY
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With

    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth / 32) + 1
 
    Dim Src As DirectDrawSurface7
    Dim rDest As RECT
    Dim dArray() As Byte, sArray() As Byte
    Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
    Dim Modo As Long
   
    Set Src = SurfaceDB.Surface(GrhData(GrhIndex).FileNum)
   
    Src.GetSurfaceDesc ddsdSrc
    InvSurface.GetSurfaceDesc ddsdDest
   
    With rDest
        .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth / 32)) * 32
        .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth / 32)) * 32
        .Bottom = .Top + 32
        .Right = .Left + 32
    End With
 
   
    Dim SrcLock As Boolean
    Dim DstLock As Boolean
     
On Error GoTo HayErrorAlpha
    
    If x < 0 Then
        x = 0
    End If
    
    If y < 0 Then
        y = 0
    End If
    
    Call Src.Lock(SourceRect, ddsdSrc, DDLOCK_WAIT, 0)
    SrcLock = True
    Call InvSurface.Lock(rDest, ddsdDest, DDLOCK_WAIT, 0)
    DstLock = True
    Call InvSurface.GetLockedArray(dArray())
    Call Src.GetLockedArray(sArray())
    
'Alpha
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left + SourceRect.Left, SourceRect.Top)), ByVal VarPtr(dArray(x + x, y)), 70, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, r, g, B)

    InvSurface.Unlock rDest
    DstLock = False
    Src.Unlock SourceRect
    SrcLock = False
Exit Sub
 
HayErrorAlpha:
    If SrcLock Then
        Src.Unlock SourceRect
    End If
    
    If DstLock Then
        InvSurface.Unlock rDest
    End If
End Sub

Public Sub SetSlot(ByVal Slot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer, ByVal GrhIndex As Integer, _
                        ByVal ObjType As eObjType, ByVal MinHit As Integer, ByVal MaxHit As Integer, _
                        ByVal MinDef As Byte, ByVal MaxDef As Byte, ByVal Valor As Long, ByVal Name As String, _
                        ByVal PuedeUsar As Boolean, ByVal Proyectil As Boolean)
    
    Select Case InventoryWindow.Tag
    
        Case frmMain.PicInv.Tag
        
            With Inv(Slot)
                .Amount = Amount
                .Name = Name
                .ObjIndex = ObjIndex
                .ObjType = ObjType
                .GrhIndex = GrhIndex
                .MinHit = MinHit
                .MaxHit = MaxHit
                .MinDef = MinDef
                .MaxDef = MaxDef
                .Valor = Valor
                .PuedeUsar = PuedeUsar
                .Proyectil = Proyectil
            End With
        
        Case frmComerciar.PicComercianteInv.Tag
    
            With NpcInv(Slot)
                .Amount = Amount
                .Name = Name
                .ObjIndex = ObjIndex
                .ObjType = ObjType
                .GrhIndex = GrhIndex
                .MinHit = MinHit
                .MaxHit = MaxHit
                .MinDef = MinDef
                .MaxDef = MaxDef
                .Valor = Valor
                .PuedeUsar = PuedeUsar
            End With
            
        Case frmBanco.PicBancoInv.Tag
            
             With Banco(Slot)
                .Amount = Amount
                .Name = Name
                .ObjIndex = ObjIndex
                .ObjType = ObjType
                .GrhIndex = GrhIndex
                .MinHit = MinHit
                .MaxHit = MaxHit
                .MinDef = MinDef
                .MaxDef = MaxDef
                .Valor = Valor
                .PuedeUsar = PuedeUsar
            End With
            
        Case frmMain.PicBelt.Tag
        
            With Belt(Slot)
                .Amount = Amount
                .Name = Name
                .ObjIndex = ObjIndex
                .GrhIndex = GrhIndex
                .Valor = Valor
            End With
            
        Case frmComerciarUsu.picInvComercio.Tag
    
    End Select

    Call DrawSlot(Slot)
    
End Sub

Public Sub UnSetSlot(ByVal Slot As Byte, Optional ByVal Draw As Boolean = True)
    
    Select Case InventoryWindow.Tag
    
        Case frmMain.PicInv.Tag
        
            With Inv(Slot)
                .Amount = 0
                .Name = vbNullString
                .ObjIndex = 0
                .ObjType = 0
                .GrhIndex = 0
                .MinHit = 0
                .MaxHit = 0
                .MinDef = 0
                .MaxDef = 0
                .Valor = 0
                .PuedeUsar = True
                .Proyectil = False
            End With
            
        Case frmMain.PicBelt.Tag
        
            With Belt(Slot)
                .Amount = 0
                .Name = vbNullString
                .ObjIndex = 0
                .GrhIndex = 0
                .Valor = 0
            End With
            
        Case frmMain.PicSpellInv.Tag

            With Spell(Slot)
                .Grh = 0
                .Nombre = vbNullString
                .ManaRequerido = 0
                .StaRequerido = 0
                .NeedStaff = 0
                .MinSkill = False
                .PuedeLanzar = False
            End With

        Case frmComerciar.PicComercianteInv.Tag
    
            With NpcInv(Slot)
                .Amount = 0
                .Name = vbNullString
                .ObjIndex = 0
                .ObjType = 0
                .GrhIndex = 0
                .MinHit = 0
                .MaxHit = 0
                .MinDef = 0
                .MaxDef = 0
                .Valor = 0
                .PuedeUsar = True
            End With

        Case frmBanco.PicBancoInv.Tag
    
            With Banco(Slot)
                .Amount = 0
                .Name = vbNullString
                .ObjIndex = 0
                .ObjType = 0
                .GrhIndex = 0
                .MinHit = 0
                .MaxHit = 0
                .MinDef = 0
                .MaxDef = 0
                .Valor = 0
                .PuedeUsar = True
            End With
           
        Case frmComerciarUsu.picInvComercio.Tag
    
    End Select
        
    If Draw Then
        Call DrawSlot(Slot)
    End If
    
End Sub

Public Sub SetSlotAmount(ByVal Slot As Byte, ByVal Amount As Integer)

    Select Case InventoryWindow.Tag
        
        Case frmMain.PicInv.Tag
            Inv(Slot).Amount = Amount
            Call DrawSlot(Slot)
            
        Case frmMain.PicBelt.Tag
            Belt(Slot).Amount = Amount
            Call DrawBeltSlot(Slot)
        
        Case frmComerciar.PicComercianteInv.Tag
            NpcInv(Slot).Amount = Amount
            Call DrawSlot(Slot)
            
        Case frmBanco.PicBancoInv.Tag
            Banco(Slot).Amount = Amount
            Call DrawSlot(Slot)
    End Select
    
End Sub
