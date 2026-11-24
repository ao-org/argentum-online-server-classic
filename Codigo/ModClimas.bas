Attribute VB_Name = "ModClimas"
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
Public Enum e_MapDirection
    e_DirectionNorth = 0
    e_DirectionEast = 1
    e_DirectionSouth = 2
    e_DirectionWest = 3
End Enum

Public Const WEATHER_ZONE_START_CHANCE_PERCENT As Integer = 5      ' Chance por minuto de que arranque una zona nueva
Public Const WEATHER_PERCENT_ROLL_MAX          As Integer = 100
Public Const WEATHER_MAX_ACTIVE_ZONES          As Integer = 3
Public Const WEATHER_ZONE_MAX_MAPS             As Integer = 9      ' Centro + 4 vecinos + 4 diagonales
Public Const WEATHER_CLOUD_LEAD_MINUTES        As Integer = 3      ' Minutos de nubes antes de que empiece a llover
Public Const WEATHER_RAIN_MIN_MINUTES          As Integer = 5
Public Const WEATHER_RAIN_MAX_MINUTES          As Integer = 15
Public Const WEATHER_FOG_MIN_INTENSITY         As Byte = 10
Public Const WEATHER_FOG_MAX_INTENSITY         As Byte = 45
Public Const WEATHER_BORDER_DEPTH              As Integer = 14     ' Tiles desde el borde donde un TileExit cuenta como salida a un mapa vecino
Public Const WEATHER_NO_DIRECTION              As Integer = -1
Public Const WEATHER_THUNDER_ROLL_MAX          As Integer = 15
Public Const WEATHER_THUNDER_ROLL_HITS         As Integer = 7      ' 7 de 15 por tick (~47%)
Public Const WEATHER_THUNDER_FIRST_WAVE        As Integer = 400
Public Const WEATHER_THUNDER_LAST_WAVE         As Integer = 403    ' El cliente solo filtra por mapa los wav 400-404
Public Const WEATHER_THUNDER_FLASH_COLOR       As Long = &HEFEECB
Public Const WEATHER_THUNDER_FLASH_MIN_MS      As Long = 80
Public Const WEATHER_THUNDER_FLASH_MAX_MS      As Long = 250
Public Const WEATHER_RAIN_START_WAVE           As Integer = 404
Public Const WEATHER_RAIN_START_FLASH_COLOR    As Long = &HD254D6
Public Const WEATHER_RAIN_START_FLASH_MS       As Long = 250

Private Type t_WeatherZone
    Active                            As Boolean
    MapCount                          As Integer
    maps(1 To WEATHER_ZONE_MAX_MAPS)  As Integer
    CloudMinutesLeft                  As Integer
    RainMinutesLeft                   As Integer
    fogIntensity                      As Byte
End Type

Private m_Zones(1 To WEATHER_MAX_ACTIVE_ZONES) As t_WeatherZone
Private m_MapNeighbors()                       As Integer
Private m_MapRaining()                         As Boolean
Private m_MapFogIntensity()                    As Byte
Private m_EligibleMaps()                       As Integer
Private m_EligibleMapCount                     As Integer
Private m_WeatherMapCount                      As Integer

Public Sub BuildWeatherNeighbors(ByVal normalMapCount As Integer)
    On Error GoTo BuildWeatherNeighbors_Err
    Dim mapIndex  As Integer
    Dim TileX     As Integer
    Dim TileY     As Integer
    Dim exitMap   As Integer
    Dim Direction As Integer
    m_WeatherMapCount = NumMaps
    ReDim m_MapNeighbors(1 To NumMaps, e_DirectionNorth To e_DirectionWest) As Integer
    ReDim m_MapRaining(1 To NumMaps) As Boolean
    ReDim m_MapFogIntensity(1 To NumMaps) As Byte
    ReDim m_EligibleMaps(1 To NumMaps) As Integer
    m_EligibleMapCount = 0
    For mapIndex = 1 To normalMapCount
        If IsWeatherEligibleMap(mapIndex) Then
            m_EligibleMapCount = m_EligibleMapCount + 1
            m_EligibleMaps(m_EligibleMapCount) = mapIndex
        End If
        For TileY = YMinMapSize To YMaxMapSize
            For TileX = XMinMapSize To XMaxMapSize
                exitMap = MapData(TileX, TileY, mapIndex).tileExit.map
                If exitMap > 0 And exitMap <= normalMapCount And exitMap <> mapIndex Then
                    Direction = GetBorderDirection(TileX, TileY)
                    If Direction <> WEATHER_NO_DIRECTION Then
                        If m_MapNeighbors(mapIndex, Direction) = 0 Then m_MapNeighbors(mapIndex, Direction) = exitMap
                    End If
                End If
            Next TileX
        Next TileY
    Next mapIndex
    Call AgregarAConsola("Servidor - Clima: " & m_EligibleMapCount & " mapas con lluvia o nieve.")
    Exit Sub
BuildWeatherNeighbors_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.BuildWeatherNeighbors", Erl)
End Sub

Public Function GetBorderDirection(ByVal TileX As Integer, ByVal TileY As Integer) As Integer
    On Error GoTo GetBorderDirection_Err
    Dim edgeCount As Integer
    GetBorderDirection = WEATHER_NO_DIRECTION
    If TileY < YMinMapSize + WEATHER_BORDER_DEPTH Then
        GetBorderDirection = e_DirectionNorth
        edgeCount = edgeCount + 1
    End If
    If TileY > YMaxMapSize - WEATHER_BORDER_DEPTH Then
        GetBorderDirection = e_DirectionSouth
        edgeCount = edgeCount + 1
    End If
    If TileX > XMaxMapSize - WEATHER_BORDER_DEPTH Then
        GetBorderDirection = e_DirectionEast
        edgeCount = edgeCount + 1
    End If
    If TileX < XMinMapSize + WEATHER_BORDER_DEPTH Then
        GetBorderDirection = e_DirectionWest
        edgeCount = edgeCount + 1
    End If
    ' Esquinas o centro del mapa: no cuentan como salida a un vecino
    If edgeCount <> 1 Then GetBorderDirection = WEATHER_NO_DIRECTION
    Exit Function
GetBorderDirection_Err:
    GetBorderDirection = WEATHER_NO_DIRECTION
    Call TraceError(Err.Number, Err.Description, "ModClimas.GetBorderDirection", Erl)
End Function

Private Function IsWeatherMapIndex(ByVal mapIndex As Integer) As Boolean
    IsWeatherMapIndex = (mapIndex >= 1 And mapIndex <= m_WeatherMapCount)
End Function

Private Function IsWeatherEligibleMap(ByVal mapIndex As Integer) As Boolean
    On Error GoTo IsWeatherEligibleMap_Err
    If mapIndex < 1 Or mapIndex > NumMaps Then Exit Function
    IsWeatherEligibleMap = (MapInfo(mapIndex).lluvia <> 0 Or MapInfo(mapIndex).Nieve <> 0)
    Exit Function
IsWeatherEligibleMap_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.IsWeatherEligibleMap", Erl)
End Function

Private Function GetMapNeighbor(ByVal mapIndex As Integer, ByVal Direction As Integer) As Integer
    If Not IsWeatherMapIndex(mapIndex) Then Exit Function
    GetMapNeighbor = m_MapNeighbors(mapIndex, Direction)
End Function

Private Function GetDiagonalNeighbor(ByVal centerMap As Integer, ByVal firstDirection As Integer, ByVal secondDirection As Integer) As Integer
    GetDiagonalNeighbor = GetMapNeighbor(GetMapNeighbor(centerMap, firstDirection), secondDirection)
    If GetDiagonalNeighbor = 0 Then
        GetDiagonalNeighbor = GetMapNeighbor(GetMapNeighbor(centerMap, secondDirection), firstDirection)
    End If
End Function

Private Function ZoneContainsMap(ByVal zoneSlot As Integer, ByVal mapIndex As Integer) As Boolean
    Dim mapSlot As Integer
    For mapSlot = 1 To m_Zones(zoneSlot).MapCount
        If m_Zones(zoneSlot).maps(mapSlot) = mapIndex Then
            ZoneContainsMap = True
            Exit Function
        End If
    Next mapSlot
End Function

Private Sub AddMapToWeatherZone(ByVal zoneSlot As Integer, ByVal mapIndex As Integer)
    On Error GoTo AddMapToWeatherZone_Err
    If Not IsWeatherEligibleMap(mapIndex) Then Exit Sub
    If ZoneContainsMap(zoneSlot, mapIndex) Then Exit Sub
    With m_Zones(zoneSlot)
        If .MapCount >= WEATHER_ZONE_MAX_MAPS Then Exit Sub
        .MapCount = .MapCount + 1
        .maps(.MapCount) = mapIndex
    End With
    Exit Sub
AddMapToWeatherZone_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.AddMapToWeatherZone", Erl)
End Sub

Private Function FindFreeWeatherZoneSlot() As Integer
    Dim zoneSlot As Integer
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If Not m_Zones(zoneSlot).Active Then
            FindFreeWeatherZoneSlot = zoneSlot
            Exit Function
        End If
    Next zoneSlot
End Function

Private Function StartWeatherZone(ByVal centerMap As Integer, ByVal cloudMinutes As Integer) As Boolean
    On Error GoTo StartWeatherZone_Err
    Dim zoneSlot As Integer
    zoneSlot = FindFreeWeatherZoneSlot()
    If zoneSlot = 0 Then Exit Function
    m_Zones(zoneSlot).MapCount = 0
    Call AddMapToWeatherZone(zoneSlot, centerMap)
    Call AddMapToWeatherZone(zoneSlot, GetMapNeighbor(centerMap, e_DirectionNorth))
    Call AddMapToWeatherZone(zoneSlot, GetMapNeighbor(centerMap, e_DirectionEast))
    Call AddMapToWeatherZone(zoneSlot, GetMapNeighbor(centerMap, e_DirectionSouth))
    Call AddMapToWeatherZone(zoneSlot, GetMapNeighbor(centerMap, e_DirectionWest))
    Call AddMapToWeatherZone(zoneSlot, GetDiagonalNeighbor(centerMap, e_DirectionNorth, e_DirectionEast))
    Call AddMapToWeatherZone(zoneSlot, GetDiagonalNeighbor(centerMap, e_DirectionNorth, e_DirectionWest))
    Call AddMapToWeatherZone(zoneSlot, GetDiagonalNeighbor(centerMap, e_DirectionSouth, e_DirectionEast))
    Call AddMapToWeatherZone(zoneSlot, GetDiagonalNeighbor(centerMap, e_DirectionSouth, e_DirectionWest))
    If m_Zones(zoneSlot).MapCount = 0 Then Exit Function
    With m_Zones(zoneSlot)
        .Active = True
        .CloudMinutesLeft = cloudMinutes
        .RainMinutesLeft = RandomNumber(WEATHER_RAIN_MIN_MINUTES, WEATHER_RAIN_MAX_MINUTES)
        .fogIntensity = RandomNumber(WEATHER_FOG_MIN_INTENSITY, WEATHER_FOG_MAX_INTENSITY)
    End With
    Call RefreshWeatherZoneMaps(zoneSlot)
    Call AgregarAConsola("Servidor - Clima: zona nueva centrada en el mapa " & centerMap & " (" & m_Zones(zoneSlot).MapCount & " mapas).")
    StartWeatherZone = True
    Exit Function
StartWeatherZone_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.StartWeatherZone", Erl)
End Function

Private Sub RefreshWeatherZoneMaps(ByVal zoneSlot As Integer)
    On Error GoTo RefreshWeatherZoneMaps_Err
    Dim mapSlot As Integer
    For mapSlot = 1 To m_Zones(zoneSlot).MapCount
        Call RefreshMapWeather(m_Zones(zoneSlot).maps(mapSlot))
    Next mapSlot
    Exit Sub
RefreshWeatherZoneMaps_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.RefreshWeatherZoneMaps", Erl)
End Sub

Private Sub RefreshMapWeather(ByVal mapIndex As Integer)
    On Error GoTo RefreshMapWeather_Err
    Dim zoneSlot     As Integer
    Dim isRaining    As Boolean
    Dim fogIntensity As Byte
    If Not IsWeatherMapIndex(mapIndex) Then Exit Sub
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active Then
            If ZoneContainsMap(zoneSlot, mapIndex) Then
                If m_Zones(zoneSlot).CloudMinutesLeft = 0 Then isRaining = True
                If m_Zones(zoneSlot).fogIntensity > fogIntensity Then fogIntensity = m_Zones(zoneSlot).fogIntensity
            End If
        End If
    Next zoneSlot
    If fogIntensity <> m_MapFogIntensity(mapIndex) Then
        m_MapFogIntensity(mapIndex) = fogIntensity
        Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageNieblandoToggle(fogIntensity))
    End If
    If isRaining <> m_MapRaining(mapIndex) Then
        m_MapRaining(mapIndex) = isRaining
        Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageRainToggle(isRaining))
        Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageNevarToggle(isRaining))
        If isRaining Then
            Call SendData(SendTarget.ToMap, mapIndex, PrepareMessagePlayWave(WEATHER_RAIN_START_WAVE, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageFlashScreen(WEATHER_RAIN_START_FLASH_COLOR, WEATHER_RAIN_START_FLASH_MS))
            Call ScheduleRainExtinguishCampfires(mapIndex)
        End If
    End If
    Exit Sub
RefreshMapWeather_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.RefreshMapWeather", Erl)
End Sub

Public Sub ProcessWeatherMinute()
    On Error GoTo ProcessWeatherMinute_Err
    Dim zoneSlot As Integer
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active Then
            If m_Zones(zoneSlot).CloudMinutesLeft > 0 Then
                m_Zones(zoneSlot).CloudMinutesLeft = m_Zones(zoneSlot).CloudMinutesLeft - 1
                ' Terminaron las nubes previas: empieza a llover
                If m_Zones(zoneSlot).CloudMinutesLeft = 0 Then Call RefreshWeatherZoneMaps(zoneSlot)
            Else
                m_Zones(zoneSlot).RainMinutesLeft = m_Zones(zoneSlot).RainMinutesLeft - 1
                If m_Zones(zoneSlot).RainMinutesLeft <= 0 Then
                    m_Zones(zoneSlot).Active = False
                    Call RefreshWeatherZoneMaps(zoneSlot)
                End If
            End If
        End If
    Next zoneSlot
    If m_EligibleMapCount > 0 Then
        If RandomNumber(1, WEATHER_PERCENT_ROLL_MAX) <= WEATHER_ZONE_START_CHANCE_PERCENT Then
            Call StartWeatherZone(m_EligibleMaps(RandomNumber(1, m_EligibleMapCount)), WEATHER_CLOUD_LEAD_MINUTES)
        End If
    End If
    Exit Sub
ProcessWeatherMinute_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.ProcessWeatherMinute", Erl)
End Sub

Public Sub ProcessWeatherThunder()
    On Error GoTo ProcessWeatherThunder_Err
    Dim zoneSlot      As Integer
    Dim mapSlot       As Integer
    Dim mapIndex      As Integer
    Dim thunderWave   As Integer
    Dim flashDuration As Long
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active Then
            If RandomNumber(1, WEATHER_THUNDER_ROLL_MAX) <= WEATHER_THUNDER_ROLL_HITS Then
                thunderWave = RandomNumber(WEATHER_THUNDER_FIRST_WAVE, WEATHER_THUNDER_LAST_WAVE)
                flashDuration = RandomNumber(WEATHER_THUNDER_FLASH_MIN_MS, WEATHER_THUNDER_FLASH_MAX_MS)
                For mapSlot = 1 To m_Zones(zoneSlot).MapCount
                    mapIndex = m_Zones(zoneSlot).maps(mapSlot)
                    If MapInfo(mapIndex).NumUsers > 0 Then
                        Call SendData(SendTarget.ToMap, mapIndex, PrepareMessagePlayWave(thunderWave, NO_3D_SOUND, NO_3D_SOUND))
                        Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageFlashScreen(WEATHER_THUNDER_FLASH_COLOR, flashDuration))
                    End If
                Next mapSlot
            End If
        End If
    Next zoneSlot
    Exit Sub
ProcessWeatherThunder_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.ProcessWeatherThunder", Erl)
End Sub

Public Sub ToggleMapWeather(ByVal mapIndex As Integer, ByVal cloudMinutes As Integer)
    On Error GoTo ToggleMapWeather_Err
    Dim zoneSlot  As Integer
    Dim endedZone As Boolean
    If Not IsWeatherMapIndex(mapIndex) Then Exit Sub
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active Then
            If ZoneContainsMap(zoneSlot, mapIndex) Then
                m_Zones(zoneSlot).Active = False
                Call RefreshWeatherZoneMaps(zoneSlot)
                endedZone = True
            End If
        End If
    Next zoneSlot
    If endedZone Then Exit Sub
    ' cloudMinutes = 0: el GM fuerza la lluvia al instante
    Call StartWeatherZone(mapIndex, cloudMinutes)
    Exit Sub
ToggleMapWeather_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.ToggleMapWeather", Erl)
End Sub

Public Sub ResetMeteo(ByVal NotifyClients As Boolean)
    On Error GoTo ResetMeteo_Err
    Dim zoneSlot  As Integer
    Dim mapIndex  As Integer
    Dim emptyZone As t_WeatherZone
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        m_Zones(zoneSlot) = emptyZone
    Next zoneSlot
    For mapIndex = 1 To m_WeatherMapCount
        If m_MapRaining(mapIndex) Then
            m_MapRaining(mapIndex) = False
            If NotifyClients Then
                Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageRainToggle(False))
                Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageNevarToggle(False))
            End If
        End If
        If m_MapFogIntensity(mapIndex) > 0 Then
            m_MapFogIntensity(mapIndex) = 0
            If NotifyClients Then Call SendData(SendTarget.ToMap, mapIndex, PrepareMessageNieblandoToggle(0))
        End If
    Next mapIndex
    frmMain.TimerMeteorologia.Enabled = True
    frmMain.Truenos.Enabled = True
    Call AgregarAConsola("Servidor > Meteorologia reseteada")
    Exit Sub
ResetMeteo_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.ResetMeteo", Erl)
End Sub

Public Function IsMapRaining(ByVal mapIndex As Integer) As Boolean
    If Not IsWeatherMapIndex(mapIndex) Then Exit Function
    IsMapRaining = m_MapRaining(mapIndex)
End Function

Public Function IsUserExposedToRain(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserExposedToRain_Err
    With UserList(UserIndex)
        If .flags.Navegando <> 0 Then Exit Function
        If MapInfo(.pos.map).lluvia = 0 Then Exit Function
        If Not IsMapRaining(.pos.map) Then Exit Function
        IsUserExposedToRain = Intemperie(UserIndex)
    End With
    Exit Function
IsUserExposedToRain_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.IsUserExposedToRain", Erl)
End Function

Public Function GetMapFogIntensity(ByVal mapIndex As Integer) As Byte
    If Not IsWeatherMapIndex(mapIndex) Then Exit Function
    GetMapFogIntensity = m_MapFogIntensity(mapIndex)
End Function

Public Function GetRainingZoneCount() As Integer
    Dim zoneSlot As Integer
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active And m_Zones(zoneSlot).CloudMinutesLeft = 0 Then
            GetRainingZoneCount = GetRainingZoneCount + 1
        End If
    Next zoneSlot
End Function

Public Function GetActiveWeatherZoneCount() As Integer
    Dim zoneSlot As Integer
    For zoneSlot = 1 To WEATHER_MAX_ACTIVE_ZONES
        If m_Zones(zoneSlot).Active Then GetActiveWeatherZoneCount = GetActiveWeatherZoneCount + 1
    Next zoneSlot
End Function

Public Sub SendMapWeatherState(ByVal UserIndex As Integer, ByVal mapIndex As Integer)
    On Error GoTo SendMapWeatherState_Err
    If IsMapRaining(mapIndex) Then
        Call WriteRainToggle(UserIndex, True)
        Call WriteNieveToggle(UserIndex, True)
    End If
    If GetMapFogIntensity(mapIndex) > 0 Then Call WriteNubesToggle(UserIndex, GetMapFogIntensity(mapIndex))
    Exit Sub
SendMapWeatherState_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.SendMapWeatherState", Erl)
End Sub

Public Sub SendWeatherTransition(ByVal UserIndex As Integer, ByVal oldMap As Integer, ByVal newMap As Integer)
    On Error GoTo SendWeatherTransition_Err
    If IsMapRaining(oldMap) <> IsMapRaining(newMap) Then
        Call WriteRainToggle(UserIndex, IsMapRaining(newMap))
        Call WriteNieveToggle(UserIndex, IsMapRaining(newMap))
    End If
    If GetMapFogIntensity(oldMap) <> GetMapFogIntensity(newMap) Then
        Call WriteNubesToggle(UserIndex, GetMapFogIntensity(newMap))
    End If
    Exit Sub
SendWeatherTransition_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.SendWeatherTransition", Erl)
End Sub
