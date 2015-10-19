Attribute VB_Name = "modCompression"
Option Explicit

Public Const GRH_SOURCE_FILE_EXT As String = ".bmp"
Public Const GRH_RESOURCE_FILE As String = "Grh.ab"
Public Const GRH_PATCH_FILE As String = "Grh.patch"

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER
    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stoRed data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, Src As Any, ByVal srclen As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0
Private Const BI_RLE8 As Long = 1
Private Const BI_RLE4 As Long = 2
Private Const BI_BITFIELDS As Long = 3
Private Const BI_JPG As Long = 4
Private Const BI_PNG As Long = 5

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal first As Long, ByVal last As Long)
'Sorts the info headers by their file name using QuickSort.

    Dim aux As INFOHEADER
    Dim Min As Long
    Dim max As Long
    Dim comp As String
    
    Min = first
    max = last
    
    comp = InfoHead((Min + max) * 0.5).strFileName
    
    Do While Min <= max
        Do While InfoHead(Min).strFileName < comp And Min < last
            Min = Min + 1
        Loop
        
        Do While InfoHead(max).strFileName > comp And max > first
            max = max - 1
        Loop
        
        If Min <= max Then
            aux = InfoHead(Min)
            InfoHead(Min) = InfoHead(max)
            InfoHead(max) = aux
            Min = Min + 1
            max = max - 1
        End If
    Loop
    
    If first < max Then
        Call Sort_Info_Headers(InfoHead, first, max)
    End If
    
    If Min < last Then
        Call Sort_Info_Headers(InfoHead, Min, last)
    End If
End Sub

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
'Searches for the specified InfoHeader

    Dim ReadingHead As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) * 0.5

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else
            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1
            End If
        End If
    Loop
End Function

Private Function Get_InfoHeader(ByRef ResourcePath As String, ByRef FileName As String, ByRef InfoHead As INFOHEADER) As Boolean
'Retrieves the InfoHead of the specified graphic file

    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    
On Error GoTo ErrHandler

    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)
        
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
            Close ResourceFile
            Exit Function
        End If
        
        'Search for it!
        If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            Get_InfoHeader = True
        End If
        
    Close ResourceFile
Exit Function

ErrHandler:
    Close ResourceFile
    
    Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & ". Razón: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

Private Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long)
'Decompresses binary data

    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Function Get_File_RawData(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte) As Boolean
'Retrieves a byte array with the compressed data from the specified file

    Dim ResourceFilePath As String
    Dim ResourceFile As Integer
    
On Error GoTo ErrHandler
    ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    
    'Size the Data array
    ReDim data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Get the data
        Get ResourceFile, InfoHead.lngFileStart, data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True
Exit Function

ErrHandler:
    Close ResourceFile
End Function

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte) As Boolean
'Extract the specific file from a resource file

On Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, data) Then
        'Decompress all data
        If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then
            Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed)
        End If
        
        Extract_File = True
    End If
Exit Function

ErrHandler:
    Call MsgBox("Error al intentar decodificar recursos. Razon: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error")
End Function

Public Sub Get_Bitmap(ByRef ResourcePath As String, ByRef FileName As String, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte)
'Retrieves bitmap file data
    
    Dim InfoHead As INFOHEADER
    Dim rawData() As Byte
    Dim offBits As Long
    Dim bitmapSize As Long
    Dim colorCount As Long
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead) Then
        'Extract the file and create the bitmap data from it.
        If Extract_File(ResourcePath, InfoHead, rawData) Then
            Call CopyMemory(offBits, rawData(10), 4)
            Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
            
            With bmpInfo.bmiHeader
                bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
                
                If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or (.biCompression <> BI_RGB And .biBitCount = 32) Then
                    If .biClrUsed < 1 Then
                        colorCount = 2 ^ .biBitCount
                    Else
                        colorCount = .biClrUsed
                    End If
                    
                    'When using bitfields on 16 or 32 bits images, bmiColors has a 3-longs mask.
                    If .biBitCount >= 16 And .biCompression = BI_BITFIELDS Then
                        colorCount = 3
                    End If
                    
                    Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), colorCount * 4)
                End If
            End With
            
            ReDim data(bitmapSize - 1) As Byte
            Call CopyMemory(data(0), rawData(offBits), bitmapSize)
        End If
    Else
        Call MsgBox("No se encontró el recurso " & FileName)
    End If
End Sub

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) / &H8
End Function
