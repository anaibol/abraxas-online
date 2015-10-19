VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TEXT_DB_ENTRY
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    Dimension As Long
    size As Long
End Type

Private mGraficos() As TEXT_DB_ENTRY
Private mD3D As D3DX8
Private device As Direct3DDevice8

Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
Private mFreeMemoryBytes As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
On Error Resume Next
    Dim i As Long
    
    For i = 1 To mCantidadGraficos
        Set mGraficos(i).texture = Nothing
    Next i
    
    Erase mGraficos
    
End Sub


Public Function GetTexture(ByVal FileName As Integer, ByRef Dimension As Long, Optional ByVal png As Byte) As Direct3DTexture8
If FileName = 0 Then
Debug.Print "0 GRH ATMPT TO BE LOADED"
Exit Function
End If
    Dim Index As Integer

    If ObtenerIndice(FileName) <> 0 Then
        With mGraficos(ObtenerIndice(FileName))
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            'Devuelvo una texture con el grafico cargado
            Dimension = .Dimension
            Set GetTexture = .texture
        End With
    Else    'Gráfico no cargado
        'Vemos si puedo agregar uno a la lista
        If mMaxEntries = mCantidadGraficos Then
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            Index = CrearGrafico(FileName, BorraMenosUsado(), , png)
            Dimension = mGraficos(Index).Dimension
            Set GetTexture = mGraficos(Index).texture
        Else
            'Agrego una textura nueva a la lista
            Index = CrearGrafico(FileName, , , png)
            If Index = 0 Then Exit Function
            Dimension = mGraficos(Index).Dimension
            Set GetTexture = mGraficos(Index).texture
        End If
    End If
End Function

Private Function ObtenerIndice(ByVal FileName As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Busqueda binaria para hallar la texture deseada
'**************************************************************
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    
    min = 1
    max = mCantidadGraficos
    Do While min <= max
        mid = (min + max) / 2
        If FileName < mGraficos(mid).FileName Then
            'El índice no existe
            If max = mid Then
                max = max - 1
            Else
                max = mid
            End If
        ElseIf FileName > mGraficos(mid).FileName Then
            'El índice no existe
            If min = mid Then
                min = min + 1
            Else
                min = mid
            End If
        Else
            ObtenerIndice = mid
            Exit Function
        End If
    Loop
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long, ByVal MaxEntries As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    
    mMaxEntries = MaxEntries
    
    If mMaxEntries < 1 Then 'por lo menos 1 gráfico
        Exit Function
    End If
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    
    Init = True
End Function



Private Function CrearGrafico(ByVal Archivo As Integer, Optional ByVal Index As Integer = -1, Optional ByVal Extraido As Boolean = False, Optional ByVal png As Byte) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
      
    If Index = -1 Then
        'Agregamos al final de la lista
        Index = mCantidadGraficos + 1
        ReDim Preserve mGraficos(1 To Index)
    End If
    
    'Call GetTextureHeader(GrhPath & Archivo & ".bmp", BMPInfo)  'para alto y ancho de la texture
            
    With mGraficos(Index)
        'Nombre
        .FileName = Archivo
        
        'Ultimo acceso
        .UltimoAcceso = GetTickCount
        
        Debug.Print DirGraficos & IIf(png, "p", "") & LTrim(Str(Archivo)) & ".png"
        Set .texture = mD3D.CreateTextureFromFileEx(device, DirGraficos & IIf(png, "p", "") & LTrim(Str(Archivo)) & ".png", _
            D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
            D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)

        .texture.GetLevelDesc 0, surface_desc
        .Dimension = surface_desc.Width
        .size = surface_desc.size
            
        mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.size
    End With
    
    'Aumentamos la cantidad de gráficos
    mCantidadGraficos = mCantidadGraficos + 1
        
    'Ordenamos la lista (al estar agregando de a un item no es mucho trabajo, y se compensa con la velocidad de busqueda)
    Call OrdenarGraficos(1, mCantidadGraficos)
    
    'Devolvemos el ìndice en que lo cargamos
    CrearGrafico = ObtenerIndice(Archivo)
Exit Function

ErrHandler:
    Debug.Print "ERROR EN GRHLOAD"
End Function

Private Sub OrdenarGraficos(ByVal primero As Integer, ByVal ultimo As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Ordenamos los gráficos por nombre usando QuickSort
'**************************************************************
    Dim min As Integer      'Primer item de la lista
    Dim max As Integer      'Ultimo item de la lista
    Dim comp As Integer     'Item usado para comparar
    Dim temp As TEXT_DB_ENTRY
    
    min = primero
    max = ultimo
    
    comp = mGraficos((min + max) / 2).FileName
    
    Do While min <= max
        Do While mGraficos(min).FileName < comp And min < ultimo
            min = min + 1
        Loop
        Do While mGraficos(max).FileName > comp And max > primero
            max = max - 1
        Loop
        If min <= max Then
            temp = mGraficos(min)
            mGraficos(min) = mGraficos(max)
            mGraficos(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    If primero < max Then OrdenarGraficos primero, max
    If min < ultimo Then OrdenarGraficos min, ultimo
End Sub

Private Function BorraMenosUsado() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Borra el grafico menos usado. Devuelve el index del borrado para ser reutilizado
'NO redimenciona la lista, estamos forzando a sacar uno para meter otro en su lugar
'**************************************************************
    Dim Valor As Long
    Dim i As Long
    
    'Inicializamos todo
    Valor = mGraficos(1).UltimoAcceso
    BorraMenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To mCantidadGraficos
        If mGraficos(i).UltimoAcceso < Valor Then
            Valor = mGraficos(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    
    'Disminuimos el contador
    mCantidadGraficos = mCantidadGraficos - 1
    
    'Borramos la texture
    Set mGraficos(BorraMenosUsado).texture = Nothing
    mGraficos(BorraMenosUsado).Dimension = 0
    mFreeMemoryBytes = mFreeMemoryBytes + mGraficos(BorraMenosUsado).size
    mGraficos(BorraMenosUsado).size = 0
    
End Function






