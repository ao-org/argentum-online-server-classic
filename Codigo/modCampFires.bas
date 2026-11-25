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

Public Const CAMPFIRE_MAX_DURATION_MS           As Long = CAMPFIRE_BASE_DURATION_MS + MAXSKILLPOINTS * CAMPFIRE_DURATION_PER_SKILL_MS
Public Const CAMPFIRE_COAL_MS                   As Long = 60000     ' 1 minuto por carbon
Public Const CAMPFIRE_PANEL_MAX_DISTANCE        As Integer = 2
Public Const COOK_MIN_UNIT_MS                   As Long = 1000

Public CookedObjByRawObj()   As Integer
Public CookSecondsByRawObj() As Long

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
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageCampfireExtinguished(X, Y))
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

Public Sub OpenCampfirePanel(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo OpenCampfirePanel_Err
    Dim campfireIndex As Integer
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
    campfireIndex = FindCampfireIndex(UserList(UserIndex).pos.map, X, Y)
    If campfireIndex = 0 Then Exit Sub
    If Not IsUserNearCampfire(UserIndex, campfireIndex) Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        Exit Sub
    End If
    UserList(UserIndex).flags.OpenCampfireId = m_Campfires(campfireIndex).campfireId
    Call CheckCookingLost(UserIndex)
    Call SendCampfireState(UserIndex, campfireIndex)
    Exit Sub
OpenCampfirePanel_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.OpenCampfirePanel", Erl)
End Sub

Public Sub CookAtOpenCampfire(ByVal UserIndex As Integer, ByVal inventorySlot As Byte, ByVal cookAmount As Integer)
    On Error GoTo CookAtOpenCampfire_Err
    Dim campfireIndex As Integer
    Dim rawObj        As Integer
    Dim survivalSkill As Integer
    campfireIndex = GetValidOpenCampfireIndex(UserIndex)
    If campfireIndex = 0 Then Exit Sub
    Call CheckCookingLost(UserIndex)
    With UserList(UserIndex)
        If .flags.CookingTotal > 0 Then
            ' Msg2319=Ya estás cocinando algo. Retirá tu comida primero.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_ALREADY_COOKING, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Exit Sub
        End If
        If inventorySlot < 1 Or inventorySlot > .CurrentInventorySlots Then Exit Sub
        rawObj = .invent.Object(inventorySlot).ObjIndex
        If rawObj < 1 Or rawObj > UBound(CookedObjByRawObj) Then
            ' Msg2318=Eso no se puede cocinar.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_NOT_COOKABLE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Exit Sub
        End If
        If CookedObjByRawObj(rawObj) = 0 Then
            ' Msg2318=Eso no se puede cocinar.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_NOT_COOKABLE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
            Exit Sub
        End If
        If cookAmount < 1 Then Exit Sub
        If cookAmount > .invent.Object(inventorySlot).Amount Then cookAmount = .invent.Object(inventorySlot).Amount
        survivalSkill = Min(max(.Stats.UserSkills(e_Skill.Supervivencia), 0), MAXSKILLPOINTS)
        .flags.CookingCampfireId = m_Campfires(campfireIndex).campfireId
        .flags.CookingRawObj = rawObj
        .flags.CookingCookedObj = CookedObjByRawObj(rawObj)
        .flags.CookingTotal = cookAmount
        .flags.CookingRetrieved = 0
        .flags.CookingStartTick = GetTickCountRaw()
        .flags.CookingUnitMs = max(COOK_MIN_UNIT_MS, CookSecondsByRawObj(rawObj) * 1000 * (1 - (survivalSkill / MAXSKILLPOINTS) * SURVIVAL_COOK_TIME_MAX_REDUCTION))
    End With
    Call QuitarUserInvItem(UserIndex, inventorySlot, cookAmount)
    Call UpdateUserInv(False, UserIndex, inventorySlot)
    Call SubirSkill(UserIndex, e_Skill.Supervivencia)
    Call SendCampfireState(UserIndex, campfireIndex)
    Exit Sub
CookAtOpenCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.CookAtOpenCampfire", Erl)
End Sub

Public Sub RetrieveCookedFood(ByVal UserIndex As Integer)
    On Error GoTo RetrieveCookedFood_Err
    Dim campfireIndex As Integer
    Dim readyAmount   As Integer
    Dim cookedItem    As t_Obj
    campfireIndex = GetValidOpenCampfireIndex(UserIndex)
    If campfireIndex = 0 Then Exit Sub
    Call CheckCookingLost(UserIndex)
    With UserList(UserIndex).flags
        If .CookingTotal = 0 Then
            ' Msg2320=Todavía no hay comida lista.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_NOTHING_READY, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .CookingCampfireId <> m_Campfires(campfireIndex).campfireId Then
            ' Msg2325=Tu comida se está cocinando en otra fogata.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_COOKING_ELSEWHERE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        readyAmount = GetCookingReadyCount(UserIndex) - .CookingRetrieved
        If readyAmount <= 0 Then
            ' Msg2320=Todavía no hay comida lista.
            Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_NOTHING_READY, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        cookedItem.ObjIndex = .CookingCookedObj
        cookedItem.Amount = readyAmount
    End With
    If Not MeterItemEnInventario(UserIndex, cookedItem) Then Exit Sub
    With UserList(UserIndex).flags
        .CookingRetrieved = .CookingRetrieved + readyAmount
        If .CookingRetrieved >= .CookingTotal Then Call ClearCooking(UserIndex)
    End With
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendCampfireState(UserIndex, campfireIndex)
    Exit Sub
RetrieveCookedFood_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.RetrieveCookedFood", Erl)
End Sub

Public Sub AddCoalToOpenCampfire(ByVal UserIndex As Integer, ByVal inventorySlot As Byte, ByVal coalAmount As Integer)
    On Error GoTo AddCoalToOpenCampfire_Err
    Dim campfireIndex   As Integer
    Dim remainingMs     As Double
    Dim maxCoalAllowed  As Long
    campfireIndex = GetValidOpenCampfireIndex(UserIndex)
    If campfireIndex = 0 Then Exit Sub
    If m_Campfires(campfireIndex).RainScheduled Then
        ' Msg2323=La lluvia está apagando la fogata, el carbón no sirve.
        Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_RAIN_NO_FUEL, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        Exit Sub
    End If
    With UserList(UserIndex)
        If inventorySlot < 1 Or inventorySlot > .CurrentInventorySlots Then Exit Sub
        If .invent.Object(inventorySlot).ObjIndex <> e_Minerales.Coal Then Exit Sub
        If coalAmount < 1 Then Exit Sub
        If coalAmount > .invent.Object(inventorySlot).Amount Then coalAmount = .invent.Object(inventorySlot).Amount
    End With
    remainingMs = GetCampfireRemainingMs(campfireIndex)
    maxCoalAllowed = Int((CAMPFIRE_MAX_DURATION_MS - remainingMs) / CAMPFIRE_COAL_MS)
    If maxCoalAllowed <= 0 Then
        ' Msg2322=La fogata no puede arder más tiempo.
        Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_FUEL_MAX, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If coalAmount > maxCoalAllowed Then coalAmount = maxCoalAllowed
    m_Campfires(campfireIndex).ExpireTick = AddMod32(m_Campfires(campfireIndex).ExpireTick, coalAmount * CAMPFIRE_COAL_MS)
    Call QuitarUserInvItem(UserIndex, inventorySlot, coalAmount)
    Call UpdateUserInv(False, UserIndex, inventorySlot)
    Call SendCampfireState(UserIndex, campfireIndex)
    Exit Sub
AddCoalToOpenCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.AddCoalToOpenCampfire", Erl)
End Sub

Public Sub ExtinguishOpenCampfire(ByVal UserIndex As Integer)
    On Error GoTo ExtinguishOpenCampfire_Err
    Dim campfireIndex As Integer
    campfireIndex = GetValidOpenCampfireIndex(UserIndex)
    If campfireIndex = 0 Then Exit Sub
    Call EraseCampfireObject(m_Campfires(campfireIndex).map, m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y)
    Call RemoveCampfireEntry(campfireIndex)
    UserList(UserIndex).flags.OpenCampfireId = 0
    ' Msg2321=Apagaste la fogata.
    Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_EXTINGUISHED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
ExtinguishOpenCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.ExtinguishOpenCampfire", Erl)
End Sub

Private Function GetValidOpenCampfireIndex(ByVal UserIndex As Integer) As Integer
    On Error GoTo GetValidOpenCampfireIndex_Err
    Dim campfireIndex As Integer
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Function
    campfireIndex = FindCampfireIndexById(UserList(UserIndex).flags.OpenCampfireId)
    If campfireIndex = 0 Then
        ' Msg2324=La fogata ya no está encendida.
        Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_NOT_LIT, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        Exit Function
    End If
    If Not IsUserNearCampfire(UserIndex, campfireIndex) Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        Exit Function
    End If
    GetValidOpenCampfireIndex = campfireIndex
    Exit Function
GetValidOpenCampfireIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.GetValidOpenCampfireIndex", Erl)
End Function

Private Function IsUserNearCampfire(ByVal UserIndex As Integer, ByVal campfireIndex As Integer) As Boolean
    On Error GoTo IsUserNearCampfire_Err
    With UserList(UserIndex).pos
        If .map <> m_Campfires(campfireIndex).map Then Exit Function
        If MapData(m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y, m_Campfires(campfireIndex).map).ObjInfo.ObjIndex <> FOGATA Then Exit Function
        IsUserNearCampfire = Abs(.X - m_Campfires(campfireIndex).X) <= CAMPFIRE_PANEL_MAX_DISTANCE And Abs(.Y - m_Campfires(campfireIndex).Y) <= CAMPFIRE_PANEL_MAX_DISTANCE
    End With
    Exit Function
IsUserNearCampfire_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.IsUserNearCampfire", Erl)
End Function

Private Function FindCampfireIndexById(ByVal campfireId As Long) As Integer
    On Error GoTo FindCampfireIndexById_Err
    Dim campfireIndex As Integer
    If campfireId = 0 Then Exit Function
    For campfireIndex = 1 To m_CampfireCount
        If m_Campfires(campfireIndex).campfireId = campfireId Then
            FindCampfireIndexById = campfireIndex
            Exit Function
        End If
    Next campfireIndex
    Exit Function
FindCampfireIndexById_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.FindCampfireIndexById", Erl)
End Function

Private Function GetCampfireRemainingMs(ByVal campfireIndex As Integer) As Double
    On Error GoTo GetCampfireRemainingMs_Err
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If DeadlinePassed(nowRaw, m_Campfires(campfireIndex).ExpireTick) Then Exit Function
    GetCampfireRemainingMs = TicksElapsed(nowRaw, m_Campfires(campfireIndex).ExpireTick)
    Exit Function
GetCampfireRemainingMs_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.GetCampfireRemainingMs", Erl)
End Function

Private Function GetCookingReadyCount(ByVal UserIndex As Integer) As Integer
    On Error GoTo GetCookingReadyCount_Err
    Dim elapsedMs As Double
    With UserList(UserIndex).flags
        If .CookingTotal = 0 Or .CookingUnitMs <= 0 Then Exit Function
        elapsedMs = TicksElapsed(.CookingStartTick, GetTickCountRaw())
        GetCookingReadyCount = Min(Int(elapsedMs / .CookingUnitMs), .CookingTotal)
    End With
    Exit Function
GetCookingReadyCount_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.GetCookingReadyCount", Erl)
End Function

Private Sub CheckCookingLost(ByVal UserIndex As Integer)
    On Error GoTo CheckCookingLost_Err
    If UserList(UserIndex).flags.CookingTotal = 0 Then Exit Sub
    If FindCampfireIndexById(UserList(UserIndex).flags.CookingCampfireId) > 0 Then Exit Sub
    Call ClearCooking(UserIndex)
    ' Msg2317=Tu comida se perdió porque la fogata se apagó.
    Call WriteLocaleMsg(UserIndex, MSG_CAMPFIRE_COOK_LOST, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
    Exit Sub
CheckCookingLost_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.CheckCookingLost", Erl)
End Sub

Private Sub ClearCooking(ByVal UserIndex As Integer)
    On Error GoTo ClearCooking_Err
    With UserList(UserIndex).flags
        .CookingCampfireId = 0
        .CookingRawObj = 0
        .CookingCookedObj = 0
        .CookingTotal = 0
        .CookingRetrieved = 0
        .CookingStartTick = 0
        .CookingUnitMs = 0
    End With
    Exit Sub
ClearCooking_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.ClearCooking", Erl)
End Sub

Private Sub SendCampfireState(ByVal UserIndex As Integer, ByVal campfireIndex As Integer)
    On Error GoTo SendCampfireState_Err
    Dim readyTotal    As Integer
    Dim secondsToNext As Long
    Dim elapsedMs     As Double
    With UserList(UserIndex).flags
        If .CookingTotal > 0 And .CookingCampfireId = m_Campfires(campfireIndex).campfireId Then
            readyTotal = GetCookingReadyCount(UserIndex)
            If readyTotal < .CookingTotal Then
                elapsedMs = TicksElapsed(.CookingStartTick, GetTickCountRaw())
                secondsToNext = -Int(-((.CookingUnitMs * (readyTotal + 1) - elapsedMs) / 1000))
            End If
            Call WriteCampfireState(UserIndex, m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y, GetCampfireRemainingMs(campfireIndex) \ 1000, .CookingRawObj, .CookingCookedObj, _
                    .CookingTotal, .CookingRetrieved, readyTotal - .CookingRetrieved, secondsToNext, .CookingUnitMs \ 1000)
        Else
            Call WriteCampfireState(UserIndex, m_Campfires(campfireIndex).X, m_Campfires(campfireIndex).Y, GetCampfireRemainingMs(campfireIndex) \ 1000, 0, 0, 0, 0, 0, 0, 0)
        End If
    End With
    Exit Sub
SendCampfireState_Err:
    Call TraceError(Err.Number, Err.Description, "modCampfires.SendCampfireState", Erl)
End Sub

