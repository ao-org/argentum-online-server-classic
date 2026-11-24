Attribute VB_Name = "Unit_Weather"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_weather() As Boolean
    Call UnitTesting.RunTest("test_weather_configuration", test_weather_configuration())
    Call UnitTesting.RunTest("test_weather_border_direction", test_weather_border_direction())
    Call UnitTesting.RunTest("test_weather_invalid_map", test_weather_invalid_map())
    Call UnitTesting.RunTest("test_weather_reset", test_weather_reset())
    test_suite_weather = True
End Function

Private Function test_weather_configuration() As Boolean
    If WEATHER_RAIN_MIN_MINUTES > WEATHER_RAIN_MAX_MINUTES Then Exit Function
    If WEATHER_FOG_MIN_INTENSITY > WEATHER_FOG_MAX_INTENSITY Then Exit Function
    If WEATHER_FOG_MIN_INTENSITY = 0 Then Exit Function
    If WEATHER_THUNDER_FIRST_WAVE > WEATHER_THUNDER_LAST_WAVE Then Exit Function
    ' El cliente solo filtra por mapa los wav 400-404, el 404 es el inicio de lluvia
    If WEATHER_THUNDER_FIRST_WAVE < 400 Or WEATHER_THUNDER_LAST_WAVE > 403 Then Exit Function
    If WEATHER_ZONE_MAX_MAPS <> 9 Then Exit Function
    If WEATHER_MAX_ACTIVE_ZONES <> 3 Then Exit Function
    test_weather_configuration = WEATHER_ZONE_START_CHANCE_PERCENT = 5
End Function

Private Function test_weather_border_direction() As Boolean
    If GetBorderDirection(50, 5) <> e_DirectionNorth Then Exit Function
    If GetBorderDirection(50, 96) <> e_DirectionSouth Then Exit Function
    If GetBorderDirection(96, 50) <> e_DirectionEast Then Exit Function
    If GetBorderDirection(5, 50) <> e_DirectionWest Then Exit Function
    ' Esquinas y centro del mapa no son salidas a un vecino
    If GetBorderDirection(5, 5) <> WEATHER_NO_DIRECTION Then Exit Function
    If GetBorderDirection(96, 96) <> WEATHER_NO_DIRECTION Then Exit Function
    test_weather_border_direction = GetBorderDirection(50, 50) = WEATHER_NO_DIRECTION
End Function

Private Function test_weather_invalid_map() As Boolean
    If IsMapRaining(0) Then Exit Function
    If IsMapRaining(-1) Then Exit Function
    test_weather_invalid_map = GetMapFogIntensity(0) = 0
End Function

Private Function test_weather_reset() As Boolean
    Call ResetMeteo(False)
    If GetRainingZoneCount() <> 0 Then Exit Function
    test_weather_reset = frmMain.TimerMeteorologia.Enabled And frmMain.Truenos.Enabled
End Function

#End If
