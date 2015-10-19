VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "diccionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clase diccionario
'basico, plain sin queso ni papa fritas


'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

Option Explicit

Private Const Max_ELEM = 100

Private Type diccElem
    Clave As String
    Def As Variant
End Type

Private p_elementos(1 To Max_ELEM) As diccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto
Private p_Cant As Integer

Public Sub Class_Initialize()
'constructor
    p_Cant = 0
End Sub

Public Sub Class_Terminate()
'destructor
    'destruir los variants?????
End Sub

Public Property Get CantElem() As Integer
    CantElem = p_Cant
End Property

Public Function AtPut(ByVal Clave As String, ByRef elem As Variant) As Boolean
    Dim i As Integer
    
    AtPut = False
    
    If LenB(Clave) = 0 Then
        Exit Function
    End If
    
    Clave = UCase$(Clave)
    
    If p_Cant = Max_ELEM Then
        AtPut = False
    Else
        For i = 1 To p_Cant
            If Clave = p_elementos(i).Clave Then
                p_elementos(i).Def = elem
                AtPut = True
                Exit For    'epa ;)
            End If
        Next i
        If Not AtPut Then
            p_Cant = p_Cant + 1
            p_elementos(p_Cant).Def = elem
            p_elementos(p_Cant).Clave = Clave
            AtPut = True
        End If
    
    End If
End Function

Public Function At(ByVal Clave As String) As Variant
    Dim i As Integer
    
    Clave = UCase$(Clave)

    At = Null
    For i = 1 To p_Cant
        If Clave = p_elementos(i).Clave Then
            At = p_elementos(i).Def
            Exit Function
        End If
    Next i

End Function

Public Function AtIndex(ByVal i As Integer) As String
    AtIndex = p_elementos(i).Clave
End Function

Public Function MayorValor(ByRef Cant As Integer) As String
    'parchecito para el AO, me da la clave con mayor valor en valor
    'y la Cantidad de claves con ese valor (por si hay empate)
    Dim i As Integer
    Dim max As Integer
    Dim Clave As String
    max = -1
    Cant = 0
    Clave = vbNullString
    For i = 1 To p_Cant
        If max <= CInt(p_elementos(i).Def) Then
            Cant = IIf(max = CInt(p_elementos(i).Def), Cant + 1, 1)
            Clave = IIf(max = CInt(p_elementos(i).Def), Clave & "," & p_elementos(i).Clave, p_elementos(i).Clave)
            max = CInt(p_elementos(i).Def)
        End If
    Next i

    MayorValor = Clave

End Function

Public Sub DumpAll()
    Dim i As Integer
    
    For i = 1 To Max_ELEM
        p_elementos(i).Clave = vbNullString
        p_elementos(i).Def = Null
    Next i
    p_Cant = 0

End Sub
