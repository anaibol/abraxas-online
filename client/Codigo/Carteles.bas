Attribute VB_Name = "modCarteles"
Option Explicit

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer

Public Sub InitCartel(Ley As String, Grh As Integer)

    If Not Cartel Then
        Leyenda = Ley
        textura = Grh
        Cartel = True
        ReDim LeyendaFormateada(0 To (Len(Ley) / (MAXLONG * 0.5)))
                    
        Dim i As Integer, k As Integer, anti As Integer
        anti = 1
        k = 0
        i = 0
        Call DarFormato(Leyenda, i, k, anti)
        i = 0
        Do While LeyendaFormateada(i) <> vbNullString And i < UBound(LeyendaFormateada)
            i = i + 1
        Loop
        ReDim Preserve LeyendaFormateada(0 To i)
    Else
        Exit Sub
    End If
    
End Sub

Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
    If anti + i <= Len(s) + 1 Then
        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
            LeyendaFormateada(k) = mid(s, anti, i + 1)
            k = k + 1
            anti = anti + i + 1
            i = 0
        Else
            i = i + 1
        End If
        Call DarFormato(s, i, k, anti)
    End If
End Function

Public Sub DibujarCartel()

    If Not Cartel Then
        Exit Sub
    End If
    
    Dim x As Integer, y As Integer
    
    x = XPosCartel + 20
    
    Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0)
    
    Dim j As Integer, desp As Integer
    
    For j = 0 To UBound(LeyendaFormateada)
        RenderText x, y + desp, LeyendaFormateada(j), &HC0FFFF, frmCharge.font
        desp = desp + (frmCharge.font.size) + 5
    Next
End Sub

