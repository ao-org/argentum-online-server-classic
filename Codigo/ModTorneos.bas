Attribute VB_Name = "ModTorneos"

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
'
Public Type t_Torneo
    HayTorneoaActivo As Boolean
    NivelMinimo As Byte
    NivelMaximo As Byte
    cupos As Byte
    costo As Long
    mago As Byte
    clerigo As Byte
    guerrero As Byte
    asesino As Byte
    bardo As Byte
    druida As Byte
    Paladin As Byte
    cazador As Byte
    Trabajador As Byte
    Pirata As Byte
    Ladron As Byte
    Bandido As Byte
    ClasesTexto As String
    participantes As Byte
    IndexParticipantes() As Integer
    Mapa As Integer
    x As Byte
    y As Byte
    nombre As String
    reglas As String
    IsStarted As Boolean
End Type

Public Torneo        As t_Torneo
Public MensajeTorneo As String
Private Const LIST_SEPARATOR As String = ", "

Public Sub IniciarTorneo()
    On Error GoTo IniciarTorneo_Err
    Dim i          As Long
    Dim inscriptos As Byte
    inscriptos = 0
    Torneo.ClasesTexto = BuildTournamentClassesText()
    If Not Torneo.HayTorneoaActivo Then
        ReDim Torneo.IndexParticipantes(1 To Torneo.cupos)
        Torneo.HayTorneoaActivo = True
    Else
        For i = 1 To Torneo.cupos
            If Torneo.IndexParticipantes(i) > 0 Then
                inscriptos = inscriptos + 1
            End If
        Next i
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_ABIERTAS_INSCRIPCIONES_CARACTERISTICAS_NIVEL_ENTRE_INSCRIPTOS, Torneo.nombre & "¬" & Torneo.NivelMinimo & "¬" & Torneo.NivelMaximo & "¬" & inscriptos & "¬" & Torneo.cupos _
            & "¬" & PonerPuntos(Torneo.costo) & "¬" & Torneo.reglas, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1674=Evento> Están abiertas las inscripciones para: ¬1: características: Nivel entre: ¬2/¬3. Inscriptos: ¬4/¬5. Precio de inscripción: ¬6 monedas de oro. Reglas: ¬7.
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_CLASES_PARTICIPANTES_ESCRIBI_PARTICIPAR_INGRESAR_EVENTO, Torneo.ClasesTexto, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1675=Evento> Clases participantes: ¬1. Escribí /PARTICIPAR para ingresar al evento.
    Exit Sub
IniciarTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.IniciarTorneo", Erl)
End Sub

Private Function BuildTournamentClassesText() As String
    On Error GoTo BuildTournamentClassesText_Err
    Dim classesText As String
    If Torneo.mago > 0 Then classesText = classesText & "Mago" & LIST_SEPARATOR
    If Torneo.druida > 0 Then classesText = classesText & "Druida" & LIST_SEPARATOR
    If Torneo.bardo > 0 Then classesText = classesText & "Bardo" & LIST_SEPARATOR
    If Torneo.clerigo > 0 Then classesText = classesText & "Clerigo" & LIST_SEPARATOR
    If Torneo.Paladin > 0 Then classesText = classesText & "Paladin" & LIST_SEPARATOR
    If Torneo.asesino > 0 Then classesText = classesText & "Asesino" & LIST_SEPARATOR
    If Torneo.Bandido > 0 Then classesText = classesText & "Bandido" & LIST_SEPARATOR
    If Torneo.cazador > 0 Then classesText = classesText & "Cazador" & LIST_SEPARATOR
    If Torneo.guerrero > 0 Then classesText = classesText & "Guerrero" & LIST_SEPARATOR
    If Torneo.Trabajador > 0 Then classesText = classesText & "Trabajador" & LIST_SEPARATOR
    If Torneo.Pirata > 0 Then classesText = classesText & "Pirata" & LIST_SEPARATOR
    If Torneo.Ladron > 0 Then classesText = classesText & "Ladron" & LIST_SEPARATOR
    If Len(classesText) > 0 Then classesText = Left$(classesText, Len(classesText) - Len(LIST_SEPARATOR))
    BuildTournamentClassesText = classesText
    Exit Function
BuildTournamentClassesText_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.BuildTournamentClassesText", Erl)
End Function

Public Sub ParticiparTorneo(ByVal UserIndex As Integer)
    On Error GoTo ParticiparTorneo_Err
    Dim IndexVacio As Byte
    With UserList(UserIndex)
        If (Not Torneo.HayTorneoaActivo) Or Torneo.IsStarted Then
            Call WriteLocaleMsg(UserIndex, MSG_HA_CERRADO_INSCRIPCION_EVENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.EnTorneo Then
            Call WriteLocaleMsg(UserIndex, MSG_ANOTADO_EVENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.EnReto Then
            Call WriteLocaleMsg(UserIndex, MSG_ENCUENTRAS_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString)
            Exit Sub
        End If
        If .Stats.GLD < Torneo.costo Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE_INGRESAR_EVENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .Stats.ELV < Torneo.NivelMinimo Then
            Call WriteLocaleMsg(UserIndex, MSG_INGRESAR_EVENTO_DEBERAS_SER_NIVEL, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, Torneo.NivelMinimo)
            Exit Sub
        End If
        If .Stats.ELV > Torneo.NivelMaximo Then
            Call WriteLocaleMsg(UserIndex, MSG_INGRESAR_EVENTO_NIVEL_NO_PUEDE_SER_MAYOR, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, Torneo.NivelMaximo)
            Exit Sub
        End If
        If Not IsClassAllowedInTournament(.clase) Then
            Call WriteLocaleMsg(UserIndex, MSG_CLASS_CANNOT_PARTICIPATE, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If MapInfo(.pos.map).Seguro = 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_SOLAMENTE_PODRAS_INGRESAR_EVENTO_ZONA_SEGURA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .Counters.Pena > 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_ANOTARTE_EVENTO_MIENTRAS_ENCARCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_ANOTARTE_EVENTO_ESTANDO_MUERTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.Comerciando Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_COMERCIANDO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.Montado = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_INGRESAR_EVENTO_ESTANDO_MONTADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If .flags.Navegando = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_NAVEGANDO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        IndexVacio = BuscarIndexFreeTorneo()
        If IndexVacio = 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_CUPOS_ESTAN_COMPLETOS, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        Torneo.IndexParticipantes(IndexVacio) = UserIndex
        Torneo.participantes = Torneo.participantes + 1
        .flags.EnTorneo = True
        .Stats.GLD = .Stats.GLD - Torneo.costo
        Call WriteUpdateGold(UserIndex)
        Call WriteLocaleMsg(UserIndex, MSG_ALREADY_REGISTERED_IN_TOURNAMENT, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_TOURNAMENT_PLAYER_JOINED, .name & Chr$(172) & (Torneo.cupos - Torneo.participantes) & Chr$(172) & Torneo.participantes & Chr$(172) & Torneo.cupos, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
        If Torneo.participantes >= Torneo.cupos Then
            Call ComenzarTorneoOk
        End If
    End With
    Exit Sub
ParticiparTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ParticiparTorneo", Erl)
End Sub

Private Function IsClassAllowedInTournament(ByVal userClass As e_Class) As Boolean
    On Error GoTo IsClassAllowedInTournament_Err
    Select Case userClass
        Case e_Class.Mage
            IsClassAllowedInTournament = Torneo.mago > 0
        Case e_Class.Druid
            IsClassAllowedInTournament = Torneo.druida > 0
        Case e_Class.Bard
            IsClassAllowedInTournament = Torneo.bardo > 0
        Case e_Class.Cleric
            IsClassAllowedInTournament = Torneo.clerigo > 0
        Case e_Class.Paladin
            IsClassAllowedInTournament = Torneo.Paladin > 0
        Case e_Class.Assasin
            IsClassAllowedInTournament = Torneo.asesino > 0
        Case e_Class.Bandit
            IsClassAllowedInTournament = Torneo.Bandido > 0
        Case e_Class.Hunter
            IsClassAllowedInTournament = Torneo.cazador > 0
        Case e_Class.Warrior
            IsClassAllowedInTournament = Torneo.guerrero > 0
        Case e_Class.Trabajador
            IsClassAllowedInTournament = Torneo.Trabajador > 0
        Case e_Class.Pirat
            IsClassAllowedInTournament = Torneo.Pirata > 0
        Case e_Class.Thief
            IsClassAllowedInTournament = Torneo.Ladron > 0
    End Select
    Exit Function
IsClassAllowedInTournament_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.IsClassAllowedInTournament", Erl)
End Function

Public Function BuscarIndexFreeTorneo() As Byte
    On Error GoTo BuscarIndexFreeTorneo_Err
    Dim i As Byte
    For i = 1 To Torneo.cupos
        If Torneo.IndexParticipantes(i) = 0 Then
            BuscarIndexFreeTorneo = i
            Exit For
        End If
    Next i
    Exit Function
BuscarIndexFreeTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.BuscarIndexFreeTorneo", Erl)
End Function

Public Sub BorrarIndexInTorneo(ByVal Index As Integer)
    On Error GoTo BorrarIndexInTorneo_Err
    Dim i As Long
    For i = 1 To Torneo.cupos
        If Torneo.IndexParticipantes(i) = Index Then
            Torneo.IndexParticipantes(i) = 0
            Torneo.participantes = Torneo.participantes - 1
            If Not Torneo.IsStarted Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_TOURNAMENT_PARTICIPANT_LEFT_SLOT_FREE, UserList(Index).name & Chr$(172) & Torneo.participantes & Chr$(172) & Torneo.cupos, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
            End If
            Exit For
        End If
    Next i
    Exit Sub
BorrarIndexInTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.BorrarIndexInTorneo", Erl)
End Sub

Public Sub ComenzarTorneoOk()
    On Error GoTo ComenzarTorneoOk_Err
    Dim nombres          As String
    Dim x                As Byte
    Dim y                As Byte
    Dim slotIndex        As Long
    Dim participantIndex As Integer
    If Not Torneo.HayTorneoaActivo Then Exit Sub
    Torneo.IsStarted = True
    For slotIndex = 1 To Torneo.cupos
        participantIndex = Torneo.IndexParticipantes(slotIndex)
        If participantIndex > 0 Then
            nombres = nombres & UserList(participantIndex).name & LIST_SEPARATOR
            x = Torneo.x
            y = Torneo.y
            Call FindLegalPos(participantIndex, Torneo.Mapa, x, y)
            Call WarpUserChar(participantIndex, Torneo.Mapa, x, y, True)
        End If
    Next slotIndex
    If Len(nombres) > 0 Then nombres = Left$(nombres, Len(nombres) - Len(LIST_SEPARATOR))
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTO_ELEGIDOS_PARTICIPAR_DAMOS_INICIO_EVENTO, nombres, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Exit Sub
ComenzarTorneoOk_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ComenzarTorneoOk", Erl)
End Sub

Public Sub ResetearTorneo()
    On Error GoTo ResetearTorneo_Err
    Dim i                As Long
    Dim participantIndex As Integer
    If Not Torneo.HayTorneoaActivo Then Exit Sub
    For i = 1 To Torneo.cupos
        participantIndex = Torneo.IndexParticipantes(i)
        If participantIndex > 0 Then
            UserList(participantIndex).flags.EnTorneo = False
            UserList(participantIndex).Stats.GLD = UserList(participantIndex).Stats.GLD + Torneo.costo
            Call WriteUpdateGold(participantIndex)
        End If
    Next i
    Torneo.HayTorneoaActivo = False
    Torneo.IsStarted = False
    Torneo.NivelMinimo = 0
    Torneo.NivelMaximo = 0
    Torneo.cupos = 0
    Torneo.costo = 0
    Torneo.mago = 0
    Torneo.clerigo = 0
    Torneo.guerrero = 0
    Torneo.asesino = 0
    Torneo.bardo = 0
    Torneo.druida = 0
    Torneo.Paladin = 0
    Torneo.cazador = 0
    Torneo.Trabajador = 0
    Torneo.Pirata = 0
    Torneo.Ladron = 0
    Torneo.Bandido = 0
    Torneo.ClasesTexto = vbNullString
    Torneo.Mapa = 0
    Torneo.x = 0
    Torneo.y = 0
    Torneo.nombre = vbNullString
    Torneo.reglas = vbNullString
    Torneo.participantes = 0
    ReDim Torneo.IndexParticipantes(1 To 1)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_TOURNAMENT_CANCELLED_BY_ADMIN, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos))
    Exit Sub
ResetearTorneo_Err:
    Call TraceError(Err.Number, Err.Description, "ModTorneos.ResetearTorneo", Erl)
End Sub
