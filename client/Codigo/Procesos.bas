Attribute VB_Name = "modProcesos"
Option Explicit

Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Const MAX_PATH As Integer = 260
 
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
 
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
"CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
 
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
 
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
 
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
 
Public Function LstPscGS() As String

On Error Resume Next
 
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    LstPscGS = vbNullString
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then
        LstPscGS = "Error"
        Exit Function
    End If
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Dim DatoP As String
    While r > 0
        If InStr(uProcess.szExeFile, ".exe") > 0 Then
            DatoP = ReadField(1, uProcess.szExeFile, Asc("."))
            Select Case DatoP
                Case "smss"
                Case "csrss"
                Case "winlogon"
                Case "services"
                Case "lsass"
                Case "svhost"
                Case "spoolsv"
                Case "cisvc"
                Case "inetinfo"
                Case "nvsvc32"
                Case "explorer"
                Case "wdfmgr"
                Case "alg"
                Case "rundll32"
                Case "soundman"
                Case "jusched"
                Case "ctfmon"
                Case "wuauclt"
                Case "svchost"
                Case "cidaemon"
                Case "wisptis"
                Case "dllhost"
                Case "wscntfy"
                Case "msdtc"
                Case Else
                    LstPscGS = LstPscGS & "|" & DatoP
            End Select
        End If
        r = ProcessNext(hSnapShot, uProcess)
    Wend
    Call CloseHandle(hSnapShot)
End Function
 
 
 

