Attribute VB_Name = "GameIni"



Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type



Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
Cabecera.CRC = Rnd * 100
Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
Dim N As Integer
Dim GameIni As tGameIni
N = FreeFile
Open App.Path & "\" & CarpetaDeInis & "\Inicio.con" For Binary As #N
Get #N, , MiCabecera

Get #N, , GameIni

Close #N
LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
Dim N As Integer
N = FreeFile
Open App.Path & "\" & CarpetaDeInis & "\Inicio.con" For Binary As #N
Put #N, , MiCabecera
GameIniConfiguration.Password = "DAMMLAMERS!"
Put #N, , GameIniConfiguration
Close #N
End Sub
Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Public Function ExisteBMP(ByVal NumeroF As Long) As Byte
'Funcion comprueba la existencia del bmp (tanto en archivo como en archivo de recursos)
    If NumeroF < 0 Or NumeroF > MAXGrH Then
        ExisteBMP = 0
        Exit Function
    End If
    If ResourceFile = 1 Or ResourceFile = 3 Then
        If FileExist(App.Path & "\" & CarpetaGraficos & "\" & Val(NumeroF) & ".bmp", vbNormal) Then
            ExisteBMP = 1 ' existe el bmp
            Exit Function
        End If
    End If
    If ResourceFile = 2 Or ResourceFile = 3 Then
        If Val(NumeroF) > ResourceF.UltimoGrafico Or NumeroF = 0 Then
            ExisteBMP = 0
            Exit Function
        End If
        If ResourceF.graficos(NumeroF).tamaño > 0 Then
            ExisteBMP = 2 ' existe en el archivo de recursos
            Exit Function
        End If
    End If

End Function
Public Sub GetTamañoBMP(ByVal fileIndex As Integer, ByRef Alto As Long, ByRef Ancho As Long, ByRef BitCount As Integer)
    Dim datos As ArchivoBMP
    Dim filePath As String
    If ExisteBMP(fileIndex) = ResourceFile And ResourceFile = 2 Then
        Call Decryptdata(fileIndex, datos)
        Ancho = datos.bmpInfo.bmiHeader.biWidth
        Alto = datos.bmpInfo.bmiHeader.biHeight
        BitCount = datos.bmpInfo.bmiHeader.biBitCount
    ElseIf ExisteBMP(fileIndex) = ResourceFile And ResourceFile = 1 Then
        filePath = App.Path & "\" & CarpetaGraficos & "\" & CStr(fileIndex) & ".bmp"
        Call surfaceDimensions(filePath, Alto, Ancho, BitCount)
    ElseIf ResourceFile = 3 Then
        If ExisteBMP(fileIndex) = 1 Then
            filePath = App.Path & "\" & CarpetaGraficos & "\" & CStr(fileIndex) & ".bmp"
            Call surfaceDimensions(filePath, Alto, Ancho, BitCount)
        ElseIf ExisteBMP(fileIndex) = 2 Then
            Call Decryptdata(fileIndex, datos)
            Ancho = datos.bmpInfo.bmiHeader.biWidth
            Alto = datos.bmpInfo.bmiHeader.biHeight
            BitCount = datos.bmpInfo.bmiHeader.biBitCount
        End If
    End If
End Sub
Public Sub Decryptdata(ByVal FileNum As Long, ByRef datos As ArchivoBMP)
 ' Censurado! :P
End Sub
Public Sub Decryptdata2(ByVal FileNum As Long, ByRef datos As ArchivoBMP)
 ' Censurado! :P
End Sub
Public Sub surfaceDimensions(ByVal Archivo As String, ByRef Height As Long, ByRef Width As Long, ByRef BitCount As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Loads the headers of a bmp file to retrieve it's dimensions at rt
'**************************************************************
    Dim handle As Integer
    Dim bmpFileHead As BITMAPFILEHEADER
    Dim bmpInfoHead As BITMAPINFOHEADER
    
    handle = FreeFile()
    Open Archivo For Binary Access Read Lock Write As handle
        Get handle, , bmpFileHead
        Get handle, , bmpInfoHead
    Close handle
    
    Height = bmpInfoHead.biHeight
    Width = bmpInfoHead.biWidth
    BitCount = bmpInfoHead.biBitCount
End Sub

Public Function FileSize(lngWidth As Long, lngHeight As Long) As Long

    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If

End Function
Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

Public Sub CalcularPosiciones(ByRef DataIndex As BodyData, ByRef Posiciones() As Position)
Dim i As Long
Dim tGrhIndex As Long
Dim graficos(1 To 4) As Integer
For i = 1 To 4
    If DataIndex.Walk(i).GrhIndex <= 0 Then DataIndex.Walk(i).GrhIndex = 1
    If Grhdata(DataIndex.Walk(i).GrhIndex).NumFrames > 1 Then
        graficos(i) = Grhdata(DataIndex.Walk(i).GrhIndex).Frames(1)
    Else
        graficos(i) = DataIndex.Walk(i).GrhIndex
    End If
Next i
For i = 1 To 4
    tGrhIndex = Grhdata(DataIndex.Walk(i).GrhIndex).Frames(1)
    If tGrhIndex <= 0 Then Exit Sub
    If i = 1 Then
        Posiciones(i).X = ((Grhdata(graficos(2)).pixelWidth + Grhdata(graficos(4)).pixelWidth + 4) / 2) - (Grhdata(graficos(1)).pixelWidth / 2)
        Posiciones(i).Y = 0
    ElseIf i = 2 Then
        Posiciones(i).X = Grhdata(graficos(4)).pixelWidth + 2
        Posiciones(i).Y = Grhdata(graficos(1)).pixelHeight + 2
    ElseIf i = 3 Then
        Posiciones(i).X = ((Grhdata(graficos(2)).pixelWidth + Grhdata(graficos(4)).pixelWidth + 4) / 2) - (Grhdata(graficos(3)).pixelWidth / 2)
        Posiciones(i).Y = Grhdata(graficos(1)).pixelHeight + Grhdata(graficos(2)).pixelHeight + 4
    ElseIf i = 4 Then
        Posiciones(i).X = 0
        Posiciones(i).Y = Grhdata(graficos(1)).pixelHeight + 2
    End If
Next i
End Sub
 
Public Function StringRecurso(ByVal Recurso As Integer) As String
    Select Case Recurso
        Case 1
            StringRecurso = "BMP"
        Case 2
            StringRecurso = "ResF"
        Case 3
            StringRecurso "BMP o ResF"
    End Select
End Function

