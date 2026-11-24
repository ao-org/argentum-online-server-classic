Attribute VB_Name = "ModUnderworld"
Option Explicit

Public UnderworldMapPool() As t_WorldPos
Public OverworldPortalPool() As t_WorldPos
Public IsUnderworldInitialized As Boolean
Private m_UnderworldLastSpawnTimestamp As Long
Private Const UNDERWORLD_BROADCAST_MSG_ID As Integer = 2174
Private Const UNDERWORLD_PORTAL_OBJ_IDX As Integer = 6357
Private Const UNDERWORLD_CENTER_MAP_NUMBER As Integer = 127
Private Const UNDERWORLD_MS_PER_MINUTE As Long = 60000
Private Const UNDERWORLD_PORTAL_EXIT_OFFSET_Y As Integer = 4
' Estado del ciclo
Private m_IsUnderworldOpen   As Boolean
Private m_PhaseStartTick     As Long
Private m_PhaseDurationMs    As Long
Private m_PhaseWarningSent   As Boolean
Private m_IsRandomCycleMode  As Boolean
' Configuracion (minutos)
Private m_OpenMinMinutes     As Long
Private m_OpenMaxMinutes     As Long
Private m_ClosedMinMinutes   As Long
Private m_ClosedMaxMinutes   As Long
Private m_CloseWarningMinutes As Long
Private m_OpenWarningMinutes  As Long

Public Sub MaybeSpawnUnderworldPortals()
    On Error GoTo MaybeSpawnUnderworldPortals_Err
    If Not IsUnderworldInitialized Then Exit Sub
    ' Modo fijado al arrancar: sin ciclo, los portales quedan abiertos siempre
    If Not m_IsRandomCycleMode Then Exit Sub
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_UnderworldLastSpawnTimestamp, nowRaw) < IntervalUnderworldSpawn Then Exit Sub
    m_UnderworldLastSpawnTimestamp = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Dim elapsedMs As Long
    elapsedMs = TicksElapsed(m_PhaseStartTick, nowRaw)
    If m_IsUnderworldOpen Then
        If elapsedMs >= m_PhaseDurationMs Then
            Call CloseUnderworld(nowRaw)
        ElseIf Not m_PhaseWarningSent Then
            If m_PhaseDurationMs - elapsedMs <= m_CloseWarningMinutes * UNDERWORLD_MS_PER_MINUTE Then
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_UNDERWORLD_CLOSING_SOON, CStr(m_CloseWarningMinutes), e_TextChannel.TEXTCHANNEL_FACTION, e_FontTypeNames.FONTTYPE_CITIZEN))
                m_PhaseWarningSent = True
            End If
        End If
    Else
        If elapsedMs >= m_PhaseDurationMs Then
            Call OpenUnderworld(nowRaw)
        ElseIf Not m_PhaseWarningSent Then
            If m_PhaseDurationMs - elapsedMs <= m_OpenWarningMinutes * UNDERWORLD_MS_PER_MINUTE Then
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_UNDERWORLD_OPENING_SOON, CStr(m_OpenWarningMinutes), e_TextChannel.TEXTCHANNEL_FACTION, e_FontTypeNames.FONTTYPE_CITIZEN))
                m_PhaseWarningSent = True
            End If
        End If
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModUnderworld.MaybeSpawnUnderworldPortals", 100)
    Exit Sub
MaybeSpawnUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.MaybeSpawnUnderworldPortals", Erl)
End Sub

Private Sub OpenUnderworld(ByVal NowTick As Long)
    On Error GoTo OpenUnderworld_Err
    Call CreateAllUnderworldPortals
    Call StartUnderworldPhase(True, NowTick)
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(UNDERWORLD_BROADCAST_MSG_ID, vbNullString, e_TextChannel.TEXTCHANNEL_FACTION, e_FontTypeNames.FONTTYPE_CITIZEN))
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.Flames, 50, 50))
    Exit Sub
OpenUnderworld_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.OpenUnderworld", Erl)
End Sub

Private Sub CloseUnderworld(ByVal NowTick As Long)
    On Error GoTo CloseUnderworld_Err
    Call StartUnderworldPhase(False, NowTick)
    Call DestroyAllUnderworldPortals
    Call KickUsersFromUnderworld
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_UNDERWORLD_CLOSED, vbNullString, e_TextChannel.TEXTCHANNEL_FACTION, e_FontTypeNames.FONTTYPE_CITIZEN))
    Exit Sub
CloseUnderworld_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.CloseUnderworld", Erl)
End Sub

Private Sub StartUnderworldPhase(ByVal IsOpenPhase As Boolean, ByVal StartTick As Long)
    On Error GoTo StartUnderworldPhase_Err
    Dim durationMinutes As Long
    If IsOpenPhase Then
        durationMinutes = RandomNumber(m_OpenMinMinutes, m_OpenMaxMinutes)
    Else
        durationMinutes = RandomNumber(m_ClosedMinMinutes, m_ClosedMaxMinutes)
    End If
    m_IsUnderworldOpen = IsOpenPhase
    m_PhaseStartTick = StartTick
    m_PhaseDurationMs = durationMinutes * UNDERWORLD_MS_PER_MINUTE
    m_PhaseWarningSent = False
    Exit Sub
StartUnderworldPhase_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.StartUnderworldPhase", Erl)
End Sub

Private Sub CreateAllUnderworldPortals()
    On Error GoTo CreateAllUnderworldPortals_Err
    Dim i As Integer
    For i = 1 To UBound(UnderworldMapPool)
        Call CreateUnderworldTp(OverworldPortalPool(i), UnderworldMapPool(i))
    Next i
    Exit Sub
CreateAllUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.CreateAllUnderworldPortals", Erl)
End Sub

Private Sub DestroyAllUnderworldPortals()
    On Error GoTo DestroyAllUnderworldPortals_Err
    Dim i As Integer
    For i = 1 To UBound(UnderworldMapPool)
        Call DestroyUnderworldTp(OverworldPortalPool(i), UnderworldMapPool(i))
    Next i
    Exit Sub
DestroyAllUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.DestroyAllUnderworldPortals", Erl)
End Sub

Public Sub KickUsersFromUnderworld()
    On Error GoTo KickUsersFromUnderworld_Err
    Dim i As Integer
    For i = 1 To UBound(UnderworldMapPool)
        Call KickUsersFromMap(UnderworldMapPool(i).map)
    Next i
    Call KickUsersFromMap(UNDERWORLD_CENTER_MAP_NUMBER)
    Exit Sub
KickUsersFromUnderworld_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.KickUsersFromUnderworld", Erl)
End Sub

Private Sub KickUsersFromMap(ByVal MapNumber As Integer)
    On Error GoTo KickUsersFromMap_Err
    If Not MapaValido(MapNumber) Then
        Call TraceError(1, MapNumber & " is invalid", "ModUnderworld.KickUsersFromMap", 0)
        Exit Sub
    End If
    Dim LoopC      As Long
    Dim UserIndex  As Integer
    Dim HomeCityId As e_City
    Dim targetX    As Long
    Dim targetY    As Long
    ' Recorrido inverso: QuitarUser corre la lista hacia atras al sacar a cada usuario
    For LoopC = ConnGroups(MapNumber).CountEntrys To 1 Step -1
        UserIndex = ConnGroups(MapNumber).UserEntrys(LoopC)
        HomeCityId = UserList(UserIndex).Hogar
        If Not IsValidCity(HomeCityId) Then HomeCityId = e_City.cUllathorpe
        ' Tile libre mas cercano al punto de la ciudad; si no hay, se usa el punto original
        If Not FindNearestFreeTile(Cities(HomeCityId).map, Cities(HomeCityId).x, Cities(HomeCityId).y, False, SPAWN_SEARCH_MAX_RADIUS, targetX, targetY) Then
            targetX = Cities(HomeCityId).x
            targetY = Cities(HomeCityId).y
        End If
        Call WarpUserChar(UserIndex, Cities(HomeCityId).map, targetX, targetY, True)
    Next LoopC
    Exit Sub
KickUsersFromMap_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.KickUsersFromMap", Erl)
End Sub

Public Sub DestroyUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    On Error GoTo DestroyUnderworldTp_Err
    Dim exitY As Integer
    exitY = Source.y - UNDERWORLD_PORTAL_EXIT_OFFSET_Y
    If Not MapaValido(Source.map) Or Not InMapBounds(Source.map, Source.x, Source.y) Or Not InMapBounds(Source.map, Source.x, exitY) Then
        Debug.Assert False
        Exit Sub
    End If
    If MapData(Source.x, Source.y, Source.Map).ObjInfo.ObjIndex <> UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
    Call EraseObj(MapData(Source.x, Source.y, Source.Map).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
    MapData(Source.x, exitY, Source.map).tileExit.map = 0
    MapData(Source.x, exitY, Source.map).tileExit.x = 0
    MapData(Source.x, exitY, Source.map).tileExit.y = 0
    Exit Sub
DestroyUnderworldTp_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.DestroyUnderworldTp", Erl)
End Sub

Public Sub CreateUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    On Error GoTo CreateUnderworldTp_Err
    Dim exitY  As Integer
    Dim Objeto As t_Obj
    exitY = Source.y - UNDERWORLD_PORTAL_EXIT_OFFSET_Y
    If Not MapaValido(Source.map) Or Not InMapBounds(Source.map, Source.x, Source.y) Or Not InMapBounds(Source.map, Source.x, exitY) Then
        Debug.Assert False
        Exit Sub
    End If
    If Not MapaValido(Dest.Map) Or Not InMapBounds(Dest.Map, Dest.x, Dest.y) Then Exit Sub
    ' Si el portal ya estaba (por ejemplo, guardado por worldsave), igual se carga el TileExit
    If MapData(Source.x, Source.y, Source.map).ObjInfo.ObjIndex <> UNDERWORLD_PORTAL_OBJ_IDX Then
        If MapData(Source.x, Source.y, Source.Map).ObjInfo.ObjIndex > 0 Then
            Call EraseObj(MapData(Source.x, Source.y, Source.Map).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
        End If
        Objeto.Amount = 1
        Objeto.ObjIndex = UNDERWORLD_PORTAL_OBJ_IDX
        Call MakeObj(Objeto, Source.Map, Source.x, Source.y)
    End If
    MapData(Source.x, exitY, Source.map).tileExit.map = Dest.map
    MapData(Source.x, exitY, Source.map).tileExit.x = Dest.x
    MapData(Source.x, exitY, Source.map).tileExit.y = Dest.y
    Exit Sub
CreateUnderworldTp_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.CreateUnderworldTp", Erl)
End Sub

Public Function IsUserIndexInsideTheUnderworld(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To UBound(UnderworldMapPool)
            If .pos.map = UnderworldMapPool(i).map Then
                IsUserIndexInsideTheUnderworld = True
                Exit Function
            End If
        Next i
            If .pos.map = UNDERWORLD_CENTER_MAP_NUMBER Then
                IsUserIndexInsideTheUnderworld = True
                Exit Function
            End If
    End With
End Function

Public Function IsUnderworldOpen() As Boolean
    IsUnderworldOpen = m_IsUnderworldOpen
End Function

Public Sub LoadUnderworldModule()
    On Error GoTo LoadUnderworldModule_Err
    IsUnderworldInitialized = False
    m_UnderworldLastSpawnTimestamp = GetTickCountRaw()
    If Not FileExist(DatPath & "UnderworldMapPool.dat", vbArchive) Then
        Call LogError("Missing file UnderworldMapPool.Dat")
        Exit Sub
    End If
    Dim IniFile           As clsIniManager
    Dim maxUnderworldMaps As Long
    Dim i                 As Integer
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "UnderworldMapPool.dat")
    MaxUnderworldMaps = val(IniFile.GetValue("INIT", "UnderworldMapPool"))
    If MaxUnderworldMaps <= 0 Then
        Call LogError("Invalid UnderworldMapPool in UnderworldMapPool.dat")
        Exit Sub
    End If
    ' Toggle "underworld": 0 = siempre abierto, 1 = ciclo aleatorio. Se fija al arrancar.
    m_IsRandomCycleMode = IsFeatureEnabled("underworld")
    If m_IsRandomCycleMode Then
        m_OpenMinMinutes = val(IniFile.GetValue("INIT", "OpenMinMinutes"))
        m_OpenMaxMinutes = val(IniFile.GetValue("INIT", "OpenMaxMinutes"))
        m_ClosedMinMinutes = val(IniFile.GetValue("INIT", "ClosedMinMinutes"))
        m_ClosedMaxMinutes = val(IniFile.GetValue("INIT", "ClosedMaxMinutes"))
        m_CloseWarningMinutes = val(IniFile.GetValue("INIT", "CloseWarningMinutes"))
        m_OpenWarningMinutes = val(IniFile.GetValue("INIT", "OpenWarningMinutes"))
        If Not IsValidMinutesRange(m_OpenMinMinutes, m_OpenMaxMinutes) Or Not IsValidMinutesRange(m_ClosedMinMinutes, m_ClosedMaxMinutes) Then
            Call LogError("Invalid open/closed minutes in UnderworldMapPool.dat")
            Exit Sub
        End If
        If m_CloseWarningMinutes <= 0 Or m_CloseWarningMinutes >= m_OpenMinMinutes Then
            Call LogError("Invalid CloseWarningMinutes in UnderworldMapPool.dat")
            Exit Sub
        End If
        If m_OpenWarningMinutes <= 0 Or m_OpenWarningMinutes >= m_ClosedMinMinutes Then
            Call LogError("Invalid OpenWarningMinutes in UnderworldMapPool.dat")
            Exit Sub
        End If
    End If
    ReDim UnderworldMapPool(1 To maxUnderworldMaps)
    ReDim OverworldPortalPool(1 To maxUnderworldMaps)
    For i = 1 To MaxUnderworldMaps
        UnderworldMapPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "DestinationMap")))
        UnderworldMapPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "DestinationX")))
        UnderworldMapPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "DestinationY")))
        OverworldPortalPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "SourceMap")))
        OverworldPortalPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "SourceX")))
        OverworldPortalPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "SourceY")))
    Next i
    Set IniFile = Nothing
    If m_IsRandomCycleMode Then
        ' Ciclo aleatorio: arranca cerrado y se limpian portales que haya dejado un worldsave
        Call DestroyAllUnderworldPortals
        Call StartUnderworldPhase(False, m_UnderworldLastSpawnTimestamp)
    Else
        ' Siempre abierto: portales fijos desde el arranque
        Call CreateAllUnderworldPortals
        m_IsUnderworldOpen = True
    End If
    IsUnderworldInitialized = True
    Exit Sub
LoadUnderworldModule_Err:
    IsUnderworldInitialized = False
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.LoadUnderworldModule", Erl)
End Sub

Private Function IsValidMinutesRange(ByVal MinMinutes As Long, ByVal MaxMinutes As Long) As Boolean
    IsValidMinutesRange = (MinMinutes > 0 And MinMinutes <= MaxMinutes)
End Function
