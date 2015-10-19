Attribute VB_Name = "modSendData"
Option Explicit

Public Enum SendTarget
    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYGuilda
    ToGuildArea
    ToRolesMasters
    ToDeadArea
    ToPartyArea
    ToHigherAdmins
    ToGMsArea
    ToUsersAreaButGMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)

On Error Resume Next
    Dim LoopC As Long
    Dim Map As Integer
    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Logged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.Logged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYGuilda
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToHigherAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
            
        Case SendTarget.ToGMsArea
            Call SendToGMsArea(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub
            
    End Select
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserAreaButIndex(ByVal UserIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 2 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If tempIndex <> UserIndex Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).Stats.Muerto Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) > 0) Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    If Not UserList(UserIndex).Guild_Id Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).Guild_Id = UserList(UserIndex).Guild_Id Or ((UserList(tempIndex).flags.Privilegios And PlayerType.Dios) And (UserList(tempIndex).flags.Privilegios And PlayerType.RoleMaster) = 0)) Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)
    
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    If UserList(UserIndex).PartyIndex = 0 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).PartyIndex = UserList(UserIndex).PartyIndex Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String)
    
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then _
                        Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
    
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = NpcList(NpcIndex).Pos.Map
    AreaX = NpcList(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = NpcList(NpcIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    AreaX = 2 ^ (AreaX / 9)
    AreaY = 2 ^ (AreaY / 9)
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next LoopC
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
    Dim LoopC As Long
    Dim Map As Integer
    Dim tempIndex As Integer
    
    Map = UserList(UserIndex).Pos.Map
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 2 Then
        Exit Sub
    End If
        
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next LoopC
End Sub

Private Sub SendToGMsArea(ByVal UserIndex As Integer, ByVal sdData As String)
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim Map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then
        Exit Sub
    End If
    
    If ConnGroups(Map).CountEntrys < 1 Then
        Exit Sub
    End If
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And PlayerType.User Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        End If
    Next LoopC
End Sub
