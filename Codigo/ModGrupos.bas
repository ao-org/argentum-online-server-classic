Attribute VB_Name = "ModGrupos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Public Grupo            As Tgrupo
Private UniqueIdCounter As Long
Private Const MAX_GROUP_MARKERS As Byte = 6

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal InvitadoIndex As Integer)
    On Error GoTo InvitarMiembro_Err
    Dim skillsNecesarios As Integer
    Dim Remitente        As t_User: Remitente = UserList(UserIndex)
    Dim Invitado         As t_User: Invitado = UserList(InvitadoIndex)
    ' Comentando linea importante abajo, solo temporalmente hasta que el sistema de grupo nuevo este implementado 3/5/2025 - ako
    ' skillsNecesarios = 15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) \ 2
    skillsNecesarios = 0
    If Remitente.Stats.UserSkills(e_Skill.liderazgo) < skillsNecesarios Then
        Call WriteLocaleMsg(UserIndex, MSG_FALTAN_PUNTOS_LIDERAZGO_LIDERAR_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, (skillsNecesarios - Remitente.Stats.UserSkills(e_Skill.liderazgo))) ' Msg2041="Te faltan ¬1 puntos en Liderazgo para liderar un grupo."
        Exit Sub
    End If
    'HarThaoS: Si invita a un gm no lo dejo
    If EsGM(InvitadoIndex) Then
        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_INVITAR_GRUPO_GM, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2042="No puedes invitar a un grupo a un GM."
        Exit Sub
    End If
    'Si es gm tampoco lo dejo
    If EsGM(UserIndex) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_GMS_PUEDEN_FORMAR_PARTE_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2043="Los GMs no pueden formar parte de un grupo."
        Exit Sub
    End If
    If Invitado.flags.SeguroParty Then
        Call WriteLocaleMsg(UserIndex, MSG_USUARIO_DEBE_DESACTIVAR_SEGURO_GRUPOS_PODER_INVITARLO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2044="El usuario debe desactivar el seguro de grupos para poder invitarlo."
        Exit Sub
    End If
    If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_INVITAR_MAS_PERSONAS_LIMITE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, CStr(UBound(Remitente.Grupo.Miembros))) ' Msg2045="No puedes invitar a mas personas. (Límite: ¬1)"
        Exit Sub
    End If
    If (Status(UserIndex) = 0 And Status(InvitadoIndex) = 1) Or (Status(UserIndex) = 1 And Status(InvitadoIndex) = 0) Or (Status(UserIndex) = 1 And Status(InvitadoIndex) = 2) Or _
            (Status(UserIndex) = 2 And Status(InvitadoIndex) = 1) Or (Status(UserIndex) = 3 And Status(InvitadoIndex) = 0) Or (Status(UserIndex) = 0 And Status(InvitadoIndex) = _
            3) Or (Status(UserIndex) = 3 And Status(InvitadoIndex) = 2) Or (Status(UserIndex) = 2 And Status(InvitadoIndex) = 3) Then
        Call WriteLocaleMsg(UserIndex, MSG_PODES_CREAR_GRUPO_PERSONAJES_DIFERENTES_FACCIONES, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2046="No podes crear un grupo con personajes de diferentes facciones."
        Exit Sub
    End If
    If (CInt(Remitente.Stats.UserSkills(e_Skill.liderazgo)) >= (15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) / 2)) Then
        ' Si el lider tiene liderazgo asignado segun su raza, se permite una diferencia de 1 nivel mas
        If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > (SvrConfig.GetValue("PartyELVwLeadership")) Then
            'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_CREAR_GRUPO_PERSONAJES_DIFERENCIA_MAS_NIVELES, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELVwLeadership"))
            Exit Sub
        End If
    Else
        If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > SvrConfig.GetValue("PartyELV") Then
            'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_CREAR_GRUPO_PERSONAJES_DIFERENCIA_MAS_NIVELES, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELV"))
            Exit Sub
        End If
    End If
    If Invitado.Grupo.EnGrupo Then
        Call WriteLocaleMsg(UserIndex, MSG_CHARACTER_ALREADY_IN_GROUP, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
    End If
    If UserList(InvitadoIndex).flags.RespondiendoPregunta = False Then
        Call WriteLocaleMsg(UserIndex, MSG_GROUP_REQUEST_SENT_WAIT_RESPONSE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Call WriteLocaleMsg(InvitadoIndex, MSG_INVITO_UNIRSE_SU_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, Remitente.name) ' Msg2049="¬1 te invitó a unirse a su grupo."
        If Not UserList(UserIndex).Grupo.EnGrupo Then
            Call SetUserRef(UserList(UserIndex).Grupo.Lider, UserIndex)
            Call SetUserRef(UserList(UserIndex).Grupo.Miembros(1), UserIndex)
            UserList(UserIndex).Grupo.CantidadMiembros = 1
        End If
        With UserList(InvitadoIndex)
            Call SetUserRef(.Grupo.PropuestaDe, UserIndex)
            .flags.pregunta = 1
        End With
        Call WritePreguntaBox(InvitadoIndex, MSG_TE_INVITO_A_UNIRSE_A_SU_GRUPO_DESEAS_UNIRTE, Remitente.name) 'Msg1595= ¬1 te invito a unirse a su grupo. ¿Deseas unirte?
        UserList(InvitadoIndex).flags.RespondiendoPregunta = True
    Else
        Call WriteLocaleMsg(UserIndex, MSG_USUARIO_TIENE_SOLICITUD_PENDIENTE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2050="El usuario tiene una solicitud pendiente."
    End If
    Exit Sub
InvitarMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.InvitarMiembro", Erl)
End Sub

Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal Indice As Byte)
    On Error GoTo EcharMiembro_Err
    Dim UserIndexEchar As Integer
    With UserList(UserIndex).Grupo
        If Not .EnGrupo Then
            Call WriteLocaleMsg(UserIndex, MSG_NINGUN_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        If .Lider.ArrayIndex <> UserIndex Then
            Call WriteLocaleMsg(UserIndex, MSG_PODES_ECHAR_USUARIOS_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        If Indice >= .CantidadMiembros Then Exit Sub
        UserIndexEchar = .Miembros(Indice + 1).ArrayIndex
        If UserIndexEchar = 0 Then Exit Sub
        If UserIndexEchar = UserIndex Then
            Call WriteLocaleMsg(UserIndex, MSG_PODES_EXPULSARTE_TI_MISMO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
    End With
    If Not RemoveMemberFromGroup(UserIndex, UserIndexEchar) Then Exit Sub
    Call WriteLocaleMsg(UserIndex, MSG_FUE_EXPULSADO_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndexEchar).name)
    Call WriteLocaleMsg(UserIndexEchar, MSG_REMOVED_FROM_GROUP, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
    Call ClearGroupMarkers(UserIndexEchar)
    Call RefreshCharStatus(UserIndexEchar)
    Call modSendData.SendData(ToIndex, UserIndexEchar, PrepareUpdateGroupInfo(UserIndexEchar))
    Call ExpelFromGroupOnlyMap(UserIndexEchar)
    Call HandleGroupAfterMemberLeft(UserIndex)
    Exit Sub
EcharMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.EcharMiembro", Erl)
End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
    On Error GoTo SalirDeGrupo_Err
    Dim LiderIndex As Integer
    If Not UserList(UserIndex).Grupo.EnGrupo Then
        Call WriteLocaleMsg(UserIndex, MSG_NINGUN_GRUPO_2057, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
    End If
    If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then
        Call ResetUserGroupData(UserIndex)
        Exit Sub
    End If
    LiderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
    If LiderIndex = UserIndex Then
        Call FinalizarGrupo(UserIndex)
        Exit Sub
    End If
    If Not RemoveMemberFromGroup(LiderIndex, UserIndex) Then Exit Sub
    Call WriteLocaleMsg(UserIndex, MSG_REMOVED_FROM_GROUP, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
    Call WriteLocaleMsg(LiderIndex, MSG_USER_LEFT_GROUP, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
    Call ClearGroupMarkers(UserIndex)
    Call RefreshCharStatus(UserIndex)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Call ExpelFromGroupOnlyMap(UserIndex)
    Call HandleGroupAfterMemberLeft(LiderIndex)
    Exit Sub
SalirDeGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupo", Erl)
End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)
    On Error GoTo SalirDeGrupoForzado_Err
    Dim LiderIndex As Integer
    If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
    If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then
        Call ResetUserGroupData(UserIndex)
        Exit Sub
    End If
    LiderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
    If LiderIndex = UserIndex Then
        Call FinalizarGrupo(UserIndex)
        Exit Sub
    End If
    If Not RemoveMemberFromGroup(LiderIndex, UserIndex) Then Exit Sub
    Call WriteLocaleMsg(LiderIndex, MSG_USER_LEFT_GROUP, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
    Call ExpelFromGroupOnlyMap(UserIndex)
    Call HandleGroupAfterMemberLeft(LiderIndex)
    Exit Sub
SalirDeGrupoForzado_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupoForzado", Erl)
End Sub

Public Sub FinalizarGrupo(ByVal LiderIndex As Integer)
    On Error GoTo FinalizarGrupo_Err
    Dim memberSlot  As Byte
    Dim memberIndex As Integer
    If Not UserList(LiderIndex).Grupo.EnGrupo Then
        Call ResetUserGroupData(LiderIndex)
        Exit Sub
    End If
    With UserList(LiderIndex).Grupo
        For memberSlot = 1 To .CantidadMiembros
            If IsValidUserRef(.Miembros(memberSlot)) Then
                memberIndex = .Miembros(memberSlot).ArrayIndex
                If memberIndex <> LiderIndex Then
                    Call DisbandGroupMember(memberIndex, False)
                End If
            End If
        Next memberSlot
    End With
    Call DisbandGroupMember(LiderIndex, True)
    Exit Sub
FinalizarGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.FinalizarGrupo", Erl)
End Sub

Private Function GetGroupSlot(ByVal LiderIndex As Integer, ByVal UserIndex As Integer) As Byte
    On Error GoTo GetGroupSlot_Err
    Dim memberSlot As Byte
    With UserList(LiderIndex).Grupo
        For memberSlot = 1 To .CantidadMiembros
            If .Miembros(memberSlot).ArrayIndex = UserIndex Then
                GetGroupSlot = memberSlot
                Exit Function
            End If
        Next memberSlot
    End With
    Exit Function
GetGroupSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.GetGroupSlot", Erl)
End Function

Private Function GetDisplaySlot(ByVal memberSlot As Byte, ByVal receiverSlot As Byte) As Byte
    On Error GoTo GetDisplaySlot_Err
    If memberSlot > receiverSlot Then
        GetDisplaySlot = memberSlot - 1
    Else
        GetDisplaySlot = memberSlot
    End If
    Exit Function
GetDisplaySlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.GetDisplaySlot", Erl)
End Function

Private Sub ClearGroupMarkers(ByVal UserIndex As Integer)
    On Error GoTo ClearGroupMarkers_Err
    Dim markerSlot As Byte
    For markerSlot = 1 To MAX_GROUP_MARKERS
        Call WriteUbicacion(UserIndex, markerSlot, 0)
    Next markerSlot
    Exit Sub
ClearGroupMarkers_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.ClearGroupMarkers", Erl)
End Sub

Private Sub RefreshGroupMarkers(ByVal LiderIndex As Integer)
    On Error GoTo RefreshGroupMarkers_Err
    Dim receiverSlot  As Byte
    Dim memberSlot    As Byte
    Dim receiverIndex As Integer
    Dim memberIndex   As Integer
    With UserList(LiderIndex).Grupo
        For receiverSlot = 1 To .CantidadMiembros
            If IsValidUserRef(.Miembros(receiverSlot)) Then
                receiverIndex = .Miembros(receiverSlot).ArrayIndex
                Call ClearGroupMarkers(receiverIndex)
                For memberSlot = 1 To .CantidadMiembros
                    If memberSlot <> receiverSlot Then
                        If IsValidUserRef(.Miembros(memberSlot)) Then
                            memberIndex = .Miembros(memberSlot).ArrayIndex
                            If UserList(memberIndex).pos.map = UserList(receiverIndex).pos.map Then
                                Call WriteUbicacion(receiverIndex, GetDisplaySlot(memberSlot, receiverSlot), memberIndex)
                            End If
                        End If
                    End If
                Next memberSlot
            End If
        Next receiverSlot
    End With
    Exit Sub
RefreshGroupMarkers_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.RefreshGroupMarkers", Erl)
End Sub

Private Sub ResetUserGroupData(ByVal UserIndex As Integer)
    On Error GoTo ResetUserGroupData_Err
    Dim memberSlot As Byte
    With UserList(UserIndex).Grupo
        .EnGrupo = False
        .id = -1
        .CantidadMiembros = 0
        Call SetUserRef(.Lider, 0)
        Call SetUserRef(.PropuestaDe, 0)
        For memberSlot = 1 To UBound(.Miembros)
            Call SetUserRef(.Miembros(memberSlot), 0)
        Next memberSlot
    End With
    Exit Sub
ResetUserGroupData_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.ResetUserGroupData", Erl)
End Sub

Private Function RemoveMemberFromGroup(ByVal LiderIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo RemoveMemberFromGroup_Err
    Dim memberSlot As Byte
    Dim shiftSlot  As Byte
    memberSlot = GetGroupSlot(LiderIndex, UserIndex)
    If memberSlot = 0 Then Exit Function
    With UserList(LiderIndex).Grupo
        For shiftSlot = memberSlot To .CantidadMiembros - 1
            .Miembros(shiftSlot) = .Miembros(shiftSlot + 1)
        Next shiftSlot
        Call ClearUserRef(.Miembros(.CantidadMiembros))
        .CantidadMiembros = .CantidadMiembros - 1
    End With
    Call ResetUserGroupData(UserIndex)
    RemoveMemberFromGroup = True
    Exit Function
RemoveMemberFromGroup_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.RemoveMemberFromGroup", Erl)
End Function

Private Sub ExpelFromGroupOnlyMap(ByVal UserIndex As Integer)
    On Error GoTo ExpelFromGroupOnlyMap_Err
    With UserList(UserIndex)
        If MapInfo(.pos.map).OnlyGroups And MapInfo(.pos.map).Salida.map <> 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.map).Salida.map, MapInfo(.pos.map).Salida.x, MapInfo(.pos.map).Salida.y, True)
        End If
    End With
    Exit Sub
ExpelFromGroupOnlyMap_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.ExpelFromGroupOnlyMap", Erl)
End Sub

Private Sub HandleGroupAfterMemberLeft(ByVal LiderIndex As Integer)
    On Error GoTo HandleGroupAfterMemberLeft_Err
    Call RefreshGroupMarkers(LiderIndex)
    If UserList(LiderIndex).Grupo.CantidadMiembros <= 1 Then
        Call WriteLocaleMsg(LiderIndex, MSG_GROUP_DISBANDED_ALL_MEMBERS_LEFT, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Call ResetUserGroupData(LiderIndex)
        Call RefreshCharStatus(LiderIndex)
        Call modSendData.SendData(ToIndex, LiderIndex, PrepareUpdateGroupInfo(LiderIndex))
        Call ExpelFromGroupOnlyMap(LiderIndex)
    Else
        Call modSendData.SendData(ToGroup, LiderIndex, PrepareUpdateGroupInfo(LiderIndex))
    End If
    Exit Sub
HandleGroupAfterMemberLeft_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.HandleGroupAfterMemberLeft", Erl)
End Sub

Private Sub DisbandGroupMember(ByVal UserIndex As Integer, ByVal IsLeader As Boolean)
    On Error GoTo DisbandGroupMember_Err
    Call ResetUserGroupData(UserIndex)
    Call ClearGroupMarkers(UserIndex)
    If IsLeader Then
        Call WriteLocaleMsg(UserIndex, MSG_DISUELTO_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
    Else
        Call WriteLocaleMsg(UserIndex, MSG_LIDER_HA_ABANDONADO_GRUPO_GRUPO_DISUELVE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
    End If
    Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Call RefreshCharStatus(UserIndex)
    Call ExpelFromGroupOnlyMap(UserIndex)
    Exit Sub
DisbandGroupMember_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.DisbandGroupMember", Erl)
End Sub

Public Sub CompartirUbicacion(ByVal UserIndex As Integer)
    On Error GoTo CompartirUbicacion_Err
    Dim LiderIndex  As Integer
    Dim ownSlot     As Byte
    Dim memberSlot  As Byte
    Dim memberIndex As Integer
    If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then Exit Sub
    LiderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
    ownSlot = GetGroupSlot(LiderIndex, UserIndex)
    If ownSlot = 0 Then Exit Sub
    With UserList(LiderIndex).Grupo
        For memberSlot = 1 To .CantidadMiembros
            If memberSlot <> ownSlot Then
                If IsValidUserRef(.Miembros(memberSlot)) Then
                    memberIndex = .Miembros(memberSlot).ArrayIndex
                    If UserList(memberIndex).pos.map = UserList(UserIndex).pos.map Then
                        Call WriteUbicacion(memberIndex, GetDisplaySlot(ownSlot, memberSlot), UserIndex)
                        Call WriteUbicacion(UserIndex, GetDisplaySlot(memberSlot, ownSlot), memberIndex)
                    Else
                        Call WriteUbicacion(memberIndex, GetDisplaySlot(ownSlot, memberSlot), 0)
                        Call WriteUbicacion(UserIndex, GetDisplaySlot(memberSlot, ownSlot), 0)
                    End If
                End If
            End If
        Next memberSlot
    End With
    Exit Sub
CompartirUbicacion_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.CompartirUbicacion", Erl)
End Sub

Public Sub BroadcastMemberPosition(ByVal UserIndex As Integer)
    On Error GoTo BroadcastMemberPosition_Err
    Dim LiderIndex  As Integer
    Dim ownSlot     As Byte
    Dim memberSlot  As Byte
    Dim memberIndex As Integer
    If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
    If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then Exit Sub
    LiderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
    ownSlot = GetGroupSlot(LiderIndex, UserIndex)
    If ownSlot = 0 Then Exit Sub
    With UserList(LiderIndex).Grupo
        For memberSlot = 1 To .CantidadMiembros
            If memberSlot <> ownSlot Then
                If IsValidUserRef(.Miembros(memberSlot)) Then
                    memberIndex = .Miembros(memberSlot).ArrayIndex
                    If UserList(memberIndex).pos.map = UserList(UserIndex).pos.map Then
                        Call WriteUbicacion(memberIndex, GetDisplaySlot(ownSlot, memberSlot), UserIndex)
                    End If
                End If
            End If
        Next memberSlot
    End With
    Exit Sub
BroadcastMemberPosition_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.BroadcastMemberPosition", Erl)
End Sub

Public Sub GroupCreateSuccess(ByVal LiderIndex As Integer)
    Call WriteLocaleMsg(LiderIndex, MSG_EL_GRUPO_HA_SIDO_CREADO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
    With UserList(LiderIndex)
        .Grupo.EnGrupo = True
        .Grupo.Id = GetNextId()
    End With
End Sub

Public Sub AddUserToGRoup(ByVal UserIndex As Integer, ByVal GroupLiderIndex As Integer)
    On Error GoTo AddUserToGRoup_Err
    Dim Index As Byte
    If UserList(GroupLiderIndex).Grupo.CantidadMiembros >= UBound(UserList(GroupLiderIndex).Grupo.Miembros) Then
        Call WriteLocaleMsg(UserIndex, MSG_SERVIDOR_SOLICITUD_GRUPO_INVALIDA_REINTENTE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Call WriteLocaleMsg(GroupLiderIndex, MSG_PUEDES_INVITAR_MAS_PERSONAS_LIMITE, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, CStr(UBound(UserList(GroupLiderIndex).Grupo.Miembros)))
        Call SetUserRef(UserList(UserIndex).Grupo.Lider, 0)
        Call SetUserRef(UserList(UserIndex).Grupo.PropuestaDe, 0)
        Exit Sub
    End If
    For Index = 1 To UserList(GroupLiderIndex).Grupo.CantidadMiembros
        If UserList(GroupLiderIndex).Grupo.Miembros(Index).ArrayIndex = UserIndex Then
            Exit Sub
        End If
    Next Index
    UserList(GroupLiderIndex).Grupo.CantidadMiembros = UserList(GroupLiderIndex).Grupo.CantidadMiembros + 1
    Call SetUserRef(UserList(GroupLiderIndex).Grupo.Miembros(UserList(GroupLiderIndex).Grupo.CantidadMiembros), UserIndex)
    Call SetUserRef(UserList(UserIndex).Grupo.Lider, GroupLiderIndex)
    UserList(UserIndex).Grupo.EnGrupo = True
    UserList(UserIndex).Grupo.Id = UserList(GroupLiderIndex).Grupo.Id
    For Index = 2 To UserList(GroupLiderIndex).Grupo.CantidadMiembros - 1
        Call WriteLocaleMsg(UserList(GroupLiderIndex).Grupo.Miembros(Index).ArrayIndex, MSG_HA_INGRESADO_AL_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
    Next Index
    Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex, MSG_HA_INGRESADO_AL_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
    Call WriteLocaleMsg(UserIndex, MSG_SIDO_ANADIDO_GRUPO, e_TextChannel.TEXTCHANNEL_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2066="¡Has sido añadido al grupo!"
    Call RefreshCharStatus(GroupLiderIndex)
    Call RefreshCharStatus(UserIndex)
    Call CompartirUbicacion(UserIndex)
    Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Exit Sub
AddUserToGRoup_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.AddUserToGRoup", Erl)
End Sub
