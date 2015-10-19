Attribute VB_Name = "modParticles"
Option Explicit

Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
    writeprivateprofilestring Main, var, value, file
End Sub
Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
    Dim sSpaces As String
    sSpaces = Space$(100)
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
