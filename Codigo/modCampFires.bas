Attribute VB_Name = "modCampFires"
Option Explicit

Public Const MAX_ACTIVE_CAMPFIRES               As Integer = 500
Public Const CAMPFIRE_BASE_DURATION_MS          As Long = 300000    ' 5 minutos
Public Const CAMPFIRE_DURATION_PER_SKILL_MS     As Long = 60000     ' 1 minuto por punto de Supervivencia
Public Const CAMPFIRE_RAIN_EXTINGUISH_MS        As Long = 5000      ' Se apaga a los 5 segundos de empezar a llover
Public Const CAMPFIRE_BASE_RADIUS               As Integer = 1
Public Const CAMPFIRE_SKILL_PER_RADIUS          As Integer = 20     ' 1 tile extra cada 20 puntos de Supervivencia
Public Const CAMPFIRE_BURN_DAMAGE_STEP_PERCENT  As Integer = 2      ' 2%, 4%, 6%... de la vida maxima por segundo encima
Public Const CAMPFIRE_BURN_MAX_SECONDS          As Integer = 50

Private Type t_ActiveCampfire
    CampfireId    As Long
    Map           As Integer
    X             As Integer
    Y             As Integer
    Radius        As Integer
    ExpireTick    As Long
    RainScheduled As Boolean
    BurnUserIndex As Integer
    BurnSeconds   As Integer
End Type

Private m_Campfires(1 To MAX_ACTIVE_CAMPFIRES) As t_ActiveCampfire
Private m_CampfireCount                        As Integer
Private m_NextCampfireId                       As Long

Public Function RegisterCampfire(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal survivalSkill As Integer) As Boolean
    On Error GoTo RegisterCampfire_Err
    Dim existingIndex As Integer
    existingIndex = FindCampfireIndex(Map, X, Y)
    If existingIndex > 0 Then Call RemoveCampfireEntry(existingIndex)
    If m_CampfireCount >= MAX_ACTIVE_CAMPFIRES Then Exit Function
    survivalSkill = Min(max(survivalSkill, 0), MAXSKILLPOINTS)
    m_NextCampfireId = m_NextCampfireId + 1
    m_CampfireCount = m_CampfireCount + 1
    With m_Campfires(m_CampfireCount)
        .CampfireId = m_NextCampfireId
        .Map = Map
        .X = X
        .Y = Y
        .Radius = CAMPFIRE_BASE_RADIUS + survivalSkill \ CAMPFIRE_SKILL_PER_RADIUS
        .ExpireTick = AddMod32(GetTickCountRaw(), CAMPFIRE_BASE_DURATION_MS + survivalSkill * CAMPFIRE_DURATION_PER_SKILL_MS)
        .RainScheduled = False
        .BurnUserIndex = 0
        .BurnSeconds = 0
    End With
    RegisterCampfire = True
    Exit Function
RegisterCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.RegisterCampfire", Erl)
End Function

Public Function IsUserInCampfireRange(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserInCampfireRange_Err
    Dim campfireIndex As Integer
    With UserList(UserIndex).pos
        For campfireIndex = 1 To m_CampfireCount
            If m_Campfires(campfireIndex).Map = .Map Then
                If Abs(m_Campfires(campfireIndex).X - .X) <= m_Campfires(campfireIndex).Radius And Abs(m_Campfires(campfireIndex).Y - .Y) <= m_Campfires(campfireIndex).Radius Then
                    If MapData(m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y, m_Campfires(campfireIndex).Map).ObjInfo.ObjIndex = FOGATA Then
                        IsUserInCampfireRange = True
                        Exit Function
                    End If
                End If
            End If
        Next campfireIndex
    End With
    Exit Function
IsUserInCampfireRange_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.IsUserInCampfireRange", Erl)
End Function

Public Sub ScheduleRainExtinguishCampfires(ByVal mapIndex As Integer)
    On Error GoTo ScheduleRainExtinguishCampfires_Err
    Dim campfireIndex As Integer
    Dim rainDeadline  As Long
    If MapInfo(mapIndex).lluvia = 0 Then Exit Sub
    rainDeadline = AddMod32(GetTickCountRaw(), CAMPFIRE_RAIN_EXTINGUISH_MS)
    For campfireIndex = 1 To m_CampfireCount
        If m_Campfires(campfireIndex).Map = mapIndex And Not m_Campfires(campfireIndex).RainScheduled Then
            m_Campfires(campfireIndex).ExpireTick = rainDeadline
            m_Campfires(campfireIndex).RainScheduled = True
        End If
    Next campfireIndex
    Exit Sub
ScheduleRainExtinguishCampfires_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.ScheduleRainExtinguishCampfires", Erl)
End Sub

Public Sub ProcessActiveCampfires()
    On Error GoTo ProcessActiveCampfires_Err
    Dim nowRaw        As Long
    Dim campfireIndex As Integer
    nowRaw = GetTickCountRaw()
    campfireIndex = 1
    Do While campfireIndex <= m_CampfireCount
        If MapData(m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y, m_Campfires(campfireIndex).Map).ObjInfo.ObjIndex <> FOGATA Then
            Call RemoveCampfireEntry(campfireIndex)
        ElseIf DeadlinePassed(nowRaw, m_Campfires(campfireIndex).ExpireTick) Then
            Call EraseCampfireObject(m_Campfires(campfireIndex).Map, m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y)
            Call RemoveCampfireEntry(campfireIndex)
        Else
            Call BurnUserOnCampfire(campfireIndex)
            campfireIndex = campfireIndex + 1
        End If
    Loop
    Exit Sub
ProcessActiveCampfires_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.ProcessActiveCampfires", Erl)
End Sub

Private Sub BurnUserOnCampfire(ByVal campfireIndex As Integer)
    On Error GoTo BurnUserOnCampfire_Err
    Dim tileUserIndex As Integer
    Dim burnDamage    As Long
    With m_Campfires(campfireIndex)
        tileUserIndex = MapData(.X, .Y, .Map).UserIndex
        If tileUserIndex <= 0 Then
            .BurnUserIndex = 0
            .BurnSeconds = 0
            Exit Sub
        End If
        If UserList(tileUserIndex).flags.Muerto = 1 Or (UserList(tileUserIndex).flags.Privilegios And e_PlayerType.User) = 0 Then
            .BurnUserIndex = 0
            .BurnSeconds = 0
            Exit Sub
        End If
        If .BurnUserIndex <> tileUserIndex Then
            .BurnUserIndex = tileUserIndex
            .BurnSeconds = 0
        End If
        If .BurnSeconds < CAMPFIRE_BURN_MAX_SECONDS Then .BurnSeconds = .BurnSeconds + 1
        burnDamage = Porcentaje(UserList(tileUserIndex).Stats.MaxHp, .BurnSeconds * CAMPFIRE_BURN_DAMAGE_STEP_PERCENT)
        If burnDamage < 1 Then burnDamage = 1
    End With
    ' Msg2314=¡Te estás quemando! Salí del fuego.
    Call WriteLocaleMsg(tileUserIndex, MSG_CAMPFIRE_BURNING, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_FIGHT)
    If UserMod.ModifyHealth(tileUserIndex, -burnDamage, 0) Then
        ' Msg516=¡Has muerto quemado!
        Call WriteLocaleMsg(tileUserIndex, MSG_MUERTO_QUEMADO, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_FIGHT)
        Call CustomScenarios.UserDie(tileUserIndex)
        Call UserMod.UserDie(tileUserIndex)
        m_Campfires(campfireIndex).BurnUserIndex = 0
        m_Campfires(campfireIndex).BurnSeconds = 0
    End If
    Exit Sub
BurnUserOnCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.BurnUserOnCampfire", Erl)
End Sub

Private Sub EraseCampfireObject(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo EraseCampfireObject_Err
    Call EraseObj(MapData(X, Y, Map).ObjInfo.Amount, Map, X, Y)
    ' Msg2313=La fogata se apagó.
    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageLocaleMsg(MSG_CAMPFIRE_WENT_OUT, vbNullString, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO))
    Exit Sub
EraseCampfireObject_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.EraseCampfireObject", Erl)
End Sub

Private Function FindCampfireIndex(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
    On Error GoTo FindCampfireIndex_Err
    Dim campfireIndex As Integer
    For campfireIndex = 1 To m_CampfireCount
        If m_Campfires(campfireIndex).Map = Map And m_Campfires(campfireIndex).X = X And m_Campfires(campfireIndex).Y = Y Then
            FindCampfireIndex = campfireIndex
            Exit Function
        End If
    Next campfireIndex
    Exit Function
FindCampfireIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.FindCampfireIndex", Erl)
End Function

Private Sub RemoveCampfireEntry(ByVal campfireIndex As Integer)
    On Error GoTo RemoveCampfireEntry_Err
    Dim emptyCampfire As t_ActiveCampfire
    If campfireIndex < 1 Or campfireIndex > m_CampfireCount Then Exit Sub
    If campfireIndex < m_CampfireCount Then m_Campfires(campfireIndex) = m_Campfires(m_CampfireCount)
    m_Campfires(m_CampfireCount) = emptyCampfire
    m_CampfireCount = m_CampfireCount - 1
    Exit Sub
RemoveCampfireEntry_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.RemoveCampfireEntry", Erl)
End Sub
