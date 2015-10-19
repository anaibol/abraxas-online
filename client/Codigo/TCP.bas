Attribute VB_Name = "ModTCP"
Option Explicit
Public Warping As Boolean
Public LlegaronAtrib As Boolean

Public Sub Login()

    Select Case EstadoLogin
        Case Normal
            Call WriteLoginChar
        Case Creado
            Call WriteLoginNewChar
        Case Recuperando
            Call WriteRecoverChar(frmConnect.Text1.Text, frmConnect.Text2.Text)
    End Select
    
    DoEvents
    
    Call FlushBuffer
End Sub
