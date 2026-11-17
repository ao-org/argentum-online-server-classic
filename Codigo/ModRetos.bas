Attribute VB_Name = "ModRetos"
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
Option Explicit

Private Const APUESTA_MAXIMA            As Long = 100000000
Private Const OBJ_POCION_ROJA           As Integer = 38
Private Const NIVEL_MINIMO_ELO          As Byte = 33
Private Const FACTOR_ELO                As Single = 0.1
Private Const RONDAS_MAXIMAS            As Byte = 3
Private Const VICTORIAS_PARA_GANAR      As Integer = 2
Private Const PUNTAJE_ABANDONO          As Integer = 123
Private Const NPC_BANQUERO_RETO         As Integer = 3
Private Const TIEMPO_DEPOSITO_ITEMS     As Integer = 60
Private Const TIEMPO_DEPOSITO_ITEMS_DBG As Integer = 20
Private Const SIN_CUENTA_REGRESIVA      As Integer = -1

Public Retos          As t_Retos
' Clave: índice del oferente. Valor: su nombre, para validar que el slot siga siendo suyo al salir de la cola
Private ListaDeEspera As New Dictionary

Public Sub CargarInfoRetos()
    Dim File As clsIniManager
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Retos.dat")
    With Retos
        .TamañoMaximoEquipo = val(File.GetValue("Retos", "MaximoEquipo"))
        .ApuestaMinima = val(File.GetValue("Retos", "ApuestaMinima"))
        .ImpuestoApuesta = val(File.GetValue("Retos", "ImpuestoApuesta"))
        .DuracionMaxima = val(File.GetValue("Retos", "DuracionMaxima"))
        #If DEBUGGING Then
            .TiempoConteo = 3
        #Else
            .TiempoConteo = val(File.GetValue("Retos", "TiempoConteo"))
        #End If
        .TotalSalas = val(File.GetValue("Salas", "Cantidad"))
        If .TotalSalas <= 0 Then Exit Sub
        ReDim .Salas(1 To .TotalSalas)
        .SalasLibres = .TotalSalas
        .AnchoSala = val(File.GetValue("Salas", "Ancho"))
        .AltoSala = val(File.GetValue("Salas", "Alto"))
        Dim Sala As Integer, SalaStr As String
        For Sala = 1 To .TotalSalas
            SalaStr = "Sala" & Sala
            With .Salas(Sala)
                .PosIzquierda.Map = val(File.GetValue(SalaStr, "Mapa"))
                .PosIzquierda.x = val(File.GetValue(SalaStr, "X"))
                .PosIzquierda.y = val(File.GetValue(SalaStr, "Y"))
                .PosDerecha.Map = .PosIzquierda.Map
                .PosDerecha.x = .PosIzquierda.x + Retos.AnchoSala - 1
                .PosDerecha.y = .PosIzquierda.y + Retos.AltoSala - 1
            End With
        Next
    End With
    Set File = Nothing
End Sub

Public Sub CrearReto(ByVal UserIndex As Integer, JugadoresStr As String, ByVal Apuesta As Long, ByVal PocionesMaximas As Integer, Optional ByVal CaenItems As Boolean = False)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
        Dim TamanoReal As Byte: TamanoReal = Retos.TamañoMaximoEquipo * 2 - 1
        If LenB(JugadoresStr) <= 0 Then Exit Sub
        Dim Jugadores() As String: Jugadores = Split(JugadoresStr, ";", TamanoReal)
        If UBound(Jugadores) > TamanoReal - 1 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        Dim nombreRepetido As String
        nombreRepetido = NombreRepetidoEnReto(.name, Jugadores)
        If LenB(nombreRepetido) > 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, nombreRepetido) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
            Exit Sub
        End If
        Dim MaxIndexEquipo As Integer: MaxIndexEquipo = UBound(Jugadores) \ 2
        If Apuesta < Retos.ApuestaMinima Or Apuesta > APUESTA_MAXIMA Then
            Call WriteLocaleMsg(UserIndex, MSG_APUESTA_MINIMA_MONEDAS_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PonerPuntos(Retos.ApuestaMinima)) ' Msg1958=La apuesta mínima es de ¬1 monedas de oro.
            Exit Sub
        End If
        If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
        If .Stats.GLD < Apuesta Then
            ' Msg588=No tienes el oro suficiente.
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            Exit Sub
        End If
        If PocionesMaximas >= 0 Then
            If TieneObjetos(OBJ_POCION_ROJA, PocionesMaximas + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_TIENES_DEMASIADAS_POCIONES_ROJAS_CANTIDAD_MAXIMA_1443, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PocionesMaximas) ' Msg1443=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        With .flags.SolicitudReto
            .Apuesta = Apuesta
            .PocionesMaximas = PocionesMaximas
            .CaenItems = CaenItems
            ReDim .Jugadores(0 To UBound(Jugadores))
            Dim i       As Integer, tIndex As t_UserReference
            Dim Equipo1 As String, Equipo2 As String
            Equipo1 = UserList(UserIndex).name
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    If EsGmChar(Jugadores(i)) Then
                        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_JUGAR_RETOS_ADMINISTRADORES, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1444=¡No puedes jugar retos con administradores!
                        Exit Sub
                    End If
                    tIndex = NameIndex(Jugadores(i))
                    If Not IsValidUserRef(tIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, Jugadores(i)) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
                        Exit Sub
                    End If
                    If Not PuedeReto(tIndex.ArrayIndex) Then
                        Call WriteLocaleMsg(UserIndex, MSG_NO_USUARIO_PUEDE_JUGAR_RETO_MOMENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(tIndex.ArrayIndex).name) ' Msg1445=El usuario ¬1 no puede jugar un reto en este momento.
                        Exit Sub
                    End If
                    .CurIndex = tIndex
                    .nombre = UserList(.CurIndex.ArrayIndex).name
                    .Aceptado = False
                    If i Mod 2 Then
                        Equipo1 = Equipo1 & IIf((i + 1) \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                    Else
                        If LenB(Equipo2) > 0 Then
                            Equipo2 = Equipo2 & IIf(i \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                        Else
                            Equipo2 = .nombre
                        End If
                    End If
                End With
            Next
            Dim Texto1 As String, Texto2 As String, Texto3 As String
            Texto1 = UserList(UserIndex).name & "(" & UserList(UserIndex).Stats.ELV & ") te invita a jugar el siguiente reto:"
            Texto2 = Equipo1 & " vs " & Equipo2 & ". Apuesta: " & PonerPuntos(Apuesta) & " monedas de oro" & IIf(CaenItems, " y los items.", ".")
            Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).name) & " para participar en el reto."
            If PocionesMaximas >= 0 Then
                Texto2 = Texto2 & " Máximo " & PocionesMaximas & " pociones rojas."
            End If
            For i = 0 To UBound(.Jugadores)
                With .Jugadores(i)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto1, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto2, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                    Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto3, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                End With
            Next
            .Estado = e_SolicitudRetoEstado.Enviada
        End With
        Call WriteLocaleMsg(UserIndex, MSG_ENVIADO_SOLICITUD_SIGUIENTE_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1446=Has enviado una solicitud para el siguiente reto:
        Call WriteConsoleMsg(UserIndex, Texto2, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call WriteLocaleMsg(UserIndex, MSG_ESCRIBE_CANCELAR_ANULAR_SOLICITUD, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos) ' Msg1447=Escribe /CANCELAR para anular la solicitud.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CrearReto", Erl)
End Sub

Public Sub AceptarReto(ByVal UserIndex As Integer, OferenteName As String)
    On Error GoTo ErrHandler
    If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
    With UserList(UserIndex)
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
        End If
    End With
    If EsGmChar(OferenteName) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_JUGAR_RETOS_ADMINISTRADORES, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1959=¡No puedes jugar retos con administradores!
        Exit Sub
    End If
    Dim Oferente As t_UserReference
    Oferente = NameIndex(OferenteName)
    If Not IsValidUserRef(Oferente) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_CONECTADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OferenteName) ' Msg1960=¬1 no está conectado.
        Exit Sub
    End If
    With UserList(Oferente.ArrayIndex).flags.SolicitudReto
        Dim JugadorIndex As Integer
        JugadorIndex = IndiceJugadorEnSolicitud(UserIndex, Oferente.ArrayIndex)
        If JugadorIndex < 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_INVITADO_NINGUN_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(Oferente.ArrayIndex).name) ' Msg1961=¬1 no te ha invitado a ningún reto o ha sido cancelado.
            Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < .Apuesta Then
            Call WriteLocaleMsg(UserIndex, MSG_NECESITAS_MONEDAS_ORO_ACEPTAR_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, PonerPuntos(.Apuesta)) ' Msg1962=Necesitas al menos ¬1 monedas de oro para aceptar este reto.
            Exit Sub
        End If
        If .PocionesMaximas >= 0 Then
            If TieneObjetos(OBJ_POCION_ROJA, .PocionesMaximas + 1, UserIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_TIENES_DEMASIADAS_POCIONES_ROJAS_CANTIDAD_MAXIMA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, .PocionesMaximas) ' Msg1963=Tienes demasiadas pociones rojas (Cantidad máxima: ¬1).
                Exit Sub
            End If
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, UserList(UserIndex).name & " ha aceptado el reto.", e_FontTypeNames.FONTTYPE_INFO)
        .Jugadores(JugadorIndex).Aceptado = True
        Call SetUserRef(.Jugadores(JugadorIndex).CurIndex, UserIndex)
        UserList(UserIndex).flags.AceptoReto = Oferente
        Call WriteLocaleMsg(UserIndex, MSG_HAS_ACEPTADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, UserList(Oferente.ArrayIndex).name) ' Msg1964=Has aceptado el reto de ¬1.
        Dim FaltanAceptar As String
        Dim i             As Integer
        For i = 0 To UBound(.Jugadores)
            If Not .Jugadores(i).Aceptado Then
                FaltanAceptar = FaltanAceptar & .Jugadores(i).nombre & " - "
            End If
        Next
        If LenB(FaltanAceptar) > 0 Then
            FaltanAceptar = Left$(FaltanAceptar, Len(FaltanAceptar) - 3)
            Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Faltan aceptar: " & FaltanAceptar, e_FontTypeNames.FONTTYPE_New_Gris)
            Exit Sub
        End If
        Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Todos los jugadores han aceptado el reto. Buscando sala...", e_FontTypeNames.FONTTYPE_New_Gris)
        Call BuscarSala(Oferente.ArrayIndex)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AceptarReto", Erl)
End Sub

Public Sub CancelarSolicitudReto(ByVal Oferente As Integer, mensaje As String)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If .Estado = e_SolicitudRetoEstado.EnCola Then
            If ListaDeEspera.Exists(Oferente) Then
                Call ListaDeEspera.Remove(Oferente)
            End If
        End If
        .Estado = e_SolicitudRetoEstado.Libre
        Dim i As Integer, tUser As t_UserReference
        ' Enviamos a los invitados
        For i = 0 To UBound(.Jugadores)
            tUser = NameIndex(.Jugadores(i).nombre)
            If IsValidUserRef(tUser) Then
                Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                Call WriteLocaleMsg(tUser.ArrayIndex, MSG_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1965=El reto ha sido cancelado.
                If .Jugadores(i).Aceptado Then
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.AceptoReto, 0)
                End If
            End If
        Next
        ' Y al oferente por separado
        Call WriteConsoleMsg(Oferente, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call WriteLocaleMsg(Oferente, MSG_RETO_SIDO_CANCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1965=El reto ha sido cancelado.
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.CancelarSolicitudReto", Erl)
End Sub

Private Sub BuscarSala(ByVal Oferente As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Retos.SalasLibres <= 0 Then
            ListaDeEspera.item(Oferente) = UserList(Oferente).name
            .Estado = e_SolicitudRetoEstado.EnCola
            Call MensajeATodosSolicitud(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        Dim Sala As Integer, SalaAleatoria As Integer
        SalaAleatoria = RandomNumber(1, Retos.SalasLibres)
        For Sala = 1 To Retos.TotalSalas
            If Not Retos.Salas(Sala).EnUso Then
                SalaAleatoria = SalaAleatoria - 1
                If SalaAleatoria = 0 Then Exit For
            End If
        Next
        Call IniciarReto(Oferente, Sala)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.BuscarSala", Erl)
End Sub

Private Sub IniciarReto(ByVal Oferente As Integer, ByVal Sala As Integer)
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        ' Última comprobación de si todos pueden entrar/pagar
        If Not TodosPuedenReto(Oferente) Then Exit Sub
        Dim Apuesta As Long, ApuestaStr As String
        Apuesta = .Apuesta
        ApuestaStr = PonerPuntos(Apuesta)
        ' Calculamos el tamaño del equipo
        Retos.Salas(Sala).TamañoEquipoIzq = UBound(.Jugadores) \ 2 + 1
        Retos.Salas(Sala).TamañoEquipoDer = Retos.Salas(Sala).TamañoEquipoIzq
        ' Reservamos espacio para los jugadores (incluyendo al oferente)
        ReDim Retos.Salas(Sala).Jugadores(0 To UBound(.Jugadores) + 1)
        ' Tiramos una moneda (50-50) y decidimos si agregar al oferente al inicio o al final de la lista
        Dim Moneda As Byte
        Moneda = RandomNumber(0, 1)
        Dim CurIndex As Integer
        If Moneda = 0 Then
            ' Agregamos al oferente al inicio (su equipo juega a la izquierda)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
            CurIndex = CurIndex + 1
        End If
        Dim i As Integer
        ' Agregamos los jugadores alternando 1 y 1 (en los índices pares está el equipo izquierdo y en los impares el derecho - el array empieza en cero)
        For i = 0 To UBound(.Jugadores)
            Retos.Salas(Sala).Jugadores(CurIndex) = .Jugadores(i).CurIndex
            CurIndex = CurIndex + 1
            ' Reset flag
            Call SetUserRef(UserList(.Jugadores(i).CurIndex.ArrayIndex).flags.AceptoReto, 0)
        Next
        If Moneda = 1 Then
            ' Agregamos al oferente al final (su equipo juega a la derecha)
            Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
        End If
        ' Reset estado de la solicitud, ya que no la necesitamos más
        .Estado = e_SolicitudRetoEstado.Libre
    End With
    With Retos.Salas(Sala)
        .EnUso = True
        .Fase = e_FaseReto.EnJuego
        .TiempoItems = 0
        .Puntaje = 0
        .Ronda = 1
        .Apuesta = Apuesta
        .TiempoRestante = Retos.DuracionMaxima
        .CaenItems = UserList(Oferente).flags.SolicitudReto.CaenItems
        Dim tUser As t_UserReference
        For i = 0 To UBound(.Jugadores)
            tUser = .Jugadores(i)
            ' Le cobramos
            UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD - Apuesta
            Call WriteUpdateGold(tUser.ArrayIndex)
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_OTORGAS_MONEDAS_ORO_POZO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, ApuestaStr) ' Msg1966=Otorgas ¬1 monedas de oro al pozo del reto.
            ' Desmontamos
            If UserList(tUser.ArrayIndex).flags.Montado <> 0 Then
                Call DoMontar(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedSaddleObjIndex), UserList(tUser.ArrayIndex).invent.EquippedSaddleSlot)
            End If
            ' Dejamos de navegar
            If UserList(tUser.ArrayIndex).flags.Nadando <> 0 Or UserList(tUser.ArrayIndex).flags.Navegando <> 0 Then
                Call DoNavega(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).invent.EquippedShipObjIndex), UserList(tUser.ArrayIndex).invent.EquippedShipSlot)
            End If
            ' Asignamos flags
            With UserList(tUser.ArrayIndex).flags
                .EnReto = True
                .EquipoReto = IIf(i Mod 2, e_EquipoReto.Derecha, e_EquipoReto.Izquierda)
                .SalaReto = Sala
                ' Guardar posición
                .LastPos = UserList(tUser.ArrayIndex).pos
            End With
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_COMENZADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1967=¡Ha comenzado el reto!
            Call WriteLocaleMsg(tUser.ArrayIndex, MSG_ADMITIR_DERROTA_ESCRIBE_ABANDONAR, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1968=Para admitir la derrota escribe /ABANDONAR.
        Next
    End With
    Retos.SalasLibres = Retos.SalasLibres - 1
    Call iniciarRonda(Sala)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarReto", Erl)
End Sub

Private Sub iniciarRonda(ByVal Sala As Integer)
    On Error GoTo iniciarRonda_Err
    Dim indiceJugador As Integer
    Dim tUser         As t_UserReference
    With Retos.Salas(Sala)
        For indiceJugador = 0 To UBound(.Jugadores)
            tUser = .Jugadores(indiceJugador)
            If IsValidUserRef(tUser) Then
                Call RevivirYLimpiar(tUser.ArrayIndex)
                ' Usando el número de ronda y el índice, decidimos el lado al que corresponde
                If (.Ronda + indiceJugador) Mod 2 = 1 Then
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                Else
                    Call WarpToLegalPos(tUser.ArrayIndex, .PosDerecha.Map, .PosDerecha.x, .PosDerecha.y, True)
                End If
                If Retos.TiempoConteo > 0 Then
                    UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                    Call WriteStopped(tUser.ArrayIndex, True)
                End If
                Call WriteLocaleMsg(tUser.ArrayIndex, MSG_COMIENZA_RONDA_N, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, .Ronda) ' Msg1969=Comienza la ronda Nº¬1
            End If
        Next
    End With
    Exit Sub
iniciarRonda_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.iniciarRonda", Erl)
End Sub

Public Sub MuereEnReto(ByVal UserIndex As Integer)
    On Error GoTo MuereEnReto_Err
    Dim Sala   As Integer
    Dim Equipo As e_EquipoReto
    Sala = UserList(UserIndex).flags.SalaReto
    Equipo = UserList(UserIndex).flags.EquipoReto
    If Sala <= 0 Then Exit Sub
    ' Solo cuentan las muertes con la pelea en curso: en la fase de items el reto ya está pagado
    If Retos.Salas(Sala).Fase <> e_FaseReto.EnJuego Then Exit Sub
    If EquipoTieneVivos(Sala, Equipo) Then Exit Sub
    ' Están todos muertos, ganó el equipo contrario
    Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    Exit Sub
MuereEnReto_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.MuereEnReto", Erl)
End Sub

Private Sub ProcesarRondaGanada(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto)
    On Error GoTo ProcesarRondaGanada_Err
    Dim indiceJugador As Integer
    Dim nombres       As String
    With Retos.Salas(Sala)
        If Equipo = e_EquipoReto.Derecha Then
            .Puntaje = .Puntaje + 1
        Else
            .Puntaje = .Puntaje - 1
        End If
        ' Si terminó la última ronda o algún equipo ya tiene las victorias necesarias
        If .Ronda >= RONDAS_MAXIMAS Or Abs(.Puntaje) >= VICTORIAS_PARA_GANAR Then
            Call FinalizarReto(Sala)
            Exit Sub
        End If
        .Ronda = .Ronda + 1
        nombres = ArmarNombresEquipo(Sala, Equipo)
        For indiceJugador = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(indiceJugador)) Then
                Call WriteLocaleMsg(.Jugadores(indiceJugador).ArrayIndex, MSG_RONDA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, nombres) ' Msg1970=Esta ronda es para ¬1.
                Call WriteConsoleMsg(.Jugadores(indiceJugador).ArrayIndex, "", e_TextChannel.TEXTCHANNEL_EVENT, 0) ' Dejamos un espacio vertical
            End If
        Next
    End With
    Call iniciarRonda(Sala)
    Exit Sub
ProcesarRondaGanada_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.ProcesarRondaGanada", Erl)
End Sub

Public Sub FinalizarReto(ByVal Sala As Integer, Optional ByVal TiempoAgotado As Boolean)
    On Error GoTo FinalizarReto_Err
    Dim indiceJugador     As Integer
    Dim tUser             As t_UserReference
    Dim oroTotal          As Long
    Dim oroPorJugador     As Long
    Dim oroStr            As String
    Dim nombresIzquierda  As String
    Dim nombresDerecha    As String
    Dim eloTotalIzquierda As Long
    Dim eloTotalDerecha   As Long
    Dim winsIzquierda     As Long
    Dim winsDerecha       As Long
    Dim eloDiff           As Long
    Dim todosAptos        As Boolean
    Dim liberarSala       As Boolean
    Dim ganador           As e_EquipoReto
    With Retos.Salas(Sala)
        ' Un reto se finaliza una sola vez: evita pagar el pozo de nuevo
        If .Fase <> e_FaseReto.EnJuego Then Exit Sub
        .Fase = e_FaseReto.Finalizando
        ' Calculamos el oro total del premio, descontando el impuesto
        OroTotal = .Apuesta * (UBound(.Jugadores) + 1)
        OroTotal = OroTotal * (1 - Retos.ImpuestoApuesta)
        todosAptos = True
        For indiceJugador = 0 To UBound(.Jugadores)
            tUser = .Jugadores(indiceJugador)
            If IsValidUserRef(tUser) Then
                todosAptos = todosAptos And (UserList(tUser.ArrayIndex).Stats.ELV >= NIVEL_MINIMO_ELO)
                If indiceJugador Mod 2 = 0 Then
                    eloTotalIzquierda = eloTotalIzquierda + UserList(tUser.ArrayIndex).Stats.ELO
                Else
                    eloTotalDerecha = eloTotalDerecha + UserList(tUser.ArrayIndex).Stats.ELO
                End If
            End If
        Next
        nombresIzquierda = ArmarNombresEquipo(Sala, e_EquipoReto.Izquierda)
        nombresDerecha = ArmarNombresEquipo(Sala, e_EquipoReto.Derecha)
        If .Puntaje = 0 Then
            ' Empate: pagamos a todos los que no abandonaron y el ELO no da bonus
            oroPorJugador = oroTotal \ (UBound(.Jugadores) + 1)
            oroStr = PonerPuntos(oroPorJugador)
            For indiceJugador = 0 To UBound(.Jugadores)
                tUser = .Jugadores(indiceJugador)
                If IsValidUserRef(tUser) Then
                    UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + oroPorJugador
                    Call WriteUpdateGold(tUser.ArrayIndex)
                    Call WriteLocaleMsg(tUser.ArrayIndex, MSG_HAS_GANADO_MONEDAS_DE_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OroStr) ' Has ganado X monedas de oro
                    If TiempoAgotado Then
                        Call WriteLocaleMsg(tUser.ArrayIndex, MSG_AGOTADO_TIEMPO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1971=Se ha agotado el tiempo del reto.
                    End If
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    Call SacarDeReto(tUser.ArrayIndex)
                    Call DevolverPosAnterior(tUser.ArrayIndex)
                End If
            Next
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VS_NINGUNO_PUDO_VENCER_RIVAL, nombresIzquierda & "¬" & nombresDerecha, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1670=Retos » ¬1 vs ¬2. Ninguno pudo vencer a su rival.
            liberarSala = True
        Else
            ganador = ObtenerGanador(Sala)
            If ganador = e_EquipoReto.Izquierda Then
                winsIzquierda = .TamañoEquipoDer
                winsDerecha = -.TamañoEquipoIzq
            Else
                winsIzquierda = -.TamañoEquipoDer
                winsDerecha = .TamañoEquipoIzq
            End If
            ' Pagamos a los ganadores que no abandonaron
            oroPorJugador = oroTotal \ ObtenerTamañoEquipo(Sala, ganador)
            oroStr = PonerPuntos(oroPorJugador)
            For indiceJugador = 0 To UBound(.Jugadores)
                tUser = .Jugadores(indiceJugador)
                If IsValidUserRef(tUser) Then
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    If TiempoAgotado Then
                        Call WriteLocaleMsg(tUser.ArrayIndex, MSG_AGOTADO_TIEMPO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1971=Se ha agotado el tiempo del reto.
                    End If
                    If UserList(tUser.ArrayIndex).flags.EquipoReto = Ganador Then
                        UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + oroPorJugador
                        Call WriteUpdateGold(tUser.ArrayIndex)
                        Call WriteLocaleMsg(tUser.ArrayIndex, MSG_HAS_GANADO_MONEDAS_DE_ORO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, OroStr) ' Has ganado X monedas de oro
                        If .CaenItems Then
                            ' Se queda en la sala para levantar los items
                            UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = SIN_CUENTA_REGRESIVA
                            Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.Map, .PosIzquierda.x, .PosIzquierda.y, True)
                        Else
                            Call SacarDeReto(tUser.ArrayIndex)
                            Call DevolverPosAnterior(tUser.ArrayIndex)
                        End If
                    Else
                        If .CaenItems Then
                            Call TirarItemsEnPos(tUser.ArrayIndex, ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x, ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y)
                        End If
                        Call SacarDeReto(tUser.ArrayIndex)
                        Call DevolverPosAnterior(tUser.ArrayIndex)
                    End If
                End If
            Next
            Dim equipoGanador  As String
            Dim equipoPerdedor As String
            equipoGanador = IIf(ganador = e_EquipoReto.Izquierda, nombresIzquierda, nombresDerecha)
            equipoPerdedor = IIf(ganador = e_EquipoReto.Izquierda, nombresDerecha, nombresIzquierda)
            If UBound(.Jugadores) > 1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_EQUIPO_VENCIO_EQUIPO_QUEDO_BOTIN_MONEDAS, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(.Apuesta), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1671=Retos » El equipo ¬1 venció al equipo ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            Else ' 1 vs 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_RETOS_VENCIO_QUEDO_BOTIN_MONEDAS_ORO, equipoGanador & "¬" & equipoPerdedor & "¬" & PonerPuntos(.Apuesta), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1672=Retos » ¬1 venció a ¬2 y se quedó con el botín de: ¬3 monedas de oro.
            End If
            If .CaenItems Then
                Call IniciarDepositoItems(Sala)
            Else
                liberarSala = True
            End If
        End If
        ' Actualizamos el ELO de cada jugador, inspirados en `Algoritmo de 400`
        ' https://en.wikipedia.org/wiki/Elo_rating_system
        For indiceJugador = 0 To UBound(.Jugadores)
            tUser = .Jugadores(indiceJugador)
            If IsValidUserRef(tUser) Then
                If todosAptos Then
                    If indiceJugador Mod 2 = 0 Then ' Jugadores en el equipo izquierdo
                        eloDiff = winsIzquierda * (eloTotalDerecha * FACTOR_ELO)
                    Else
                        eloDiff = winsDerecha * (eloTotalIzquierda * FACTOR_ELO)
                    End If
                    If eloDiff > 0 Then
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_HAS_GANADO_PUNTOS_ELO, Abs(eloDiff), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1695=Has ganado ¬1 puntos de ELO!
                    Else
                        If UserList(tUser.ArrayIndex).Stats.ELO < Abs(eloDiff) Then
                            eloDiff = -UserList(tUser.ArrayIndex).Stats.ELO
                        End If
                        Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_HAS_PERDIDO_PUNTOS_ELO, Abs(eloDiff), e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1696=Has perdido ¬1 puntos de ELO!
                    End If
                    UserList(tUser.ArrayIndex).Stats.ELO = UserList(tUser.ArrayIndex).Stats.ELO + eloDiff
                Else
                    Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageLocaleMsg(MSG_PARTICIPANTE_RETO_TIENE_NIVEL_MENOR_ELO_PERMANECE, vbNullString, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)) 'Msg1697=Al menos un participante del reto tiene nivel menor a 35, tu ELO permanece igual.
                End If
            End If
        Next
    End With
    ' Liberamos al final y fuera del With: la cola puede arrancar otro reto en esta misma sala
    If liberarSala Then
        Call SalaLiberada(Sala)
    End If
    Exit Sub
FinalizarReto_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.FinalizarReto", Erl)
    ' Si falló a mitad de camino, no dejamos la sala trabada (la fase ya impide pagar de nuevo)
    If Retos.Salas(Sala).Fase = e_FaseReto.Finalizando Then
        Call SalaLiberada(Sala)
    End If
End Sub

Public Sub TirarItemsEnPos(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo TirarItemsEnPos_Err
    Dim i         As Byte
    Dim NuevaPos  As t_WorldPos
    Dim MiObj     As t_Obj
    Dim ItemIndex As Integer
    Dim posItems  As t_WorldPos
    With UserList(UserIndex)
        posItems.Map = .pos.Map
        posItems.x = x
        posItems.y = y
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And Not ItemNewbieProtegidoAlMorir(UserIndex, ItemIndex) Then
                    NuevaPos.x = 0
                    NuevaPos.y = 0
                    MiObj.amount = .invent.Object(i).amount
                    MiObj.ObjIndex = ItemIndex
                    MiObj.ElementalTags = .invent.Object(i).ElementalTags
                    Call Tilelibre(posItems, NuevaPos, MiObj, True, True, False)
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                        Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                        '  Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                    Else
                        Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
                        Call UpdateUserInv(False, UserIndex, i)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
TirarItemsEnPos_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.TirarItemsEnPos", Erl)
End Sub

Public Sub IniciarDepositoItems(ByVal Sala As Integer)
    On Error GoTo IniciarDepositoItems_Err
    Dim indiceJugador As Integer
    Dim ganador       As e_EquipoReto
    Dim posBanquero   As t_WorldPos
    With Retos.Salas(Sala)
        ' Primero la fase y el tiempo: si algo falla después, el timer igual cierra la sala
        .Fase = e_FaseReto.DepositoItems
        #If DEBUGGING Then
            .TiempoItems = TIEMPO_DEPOSITO_ITEMS_DBG
        #Else
            .TiempoItems = TIEMPO_DEPOSITO_ITEMS
        #End If
        ganador = ObtenerGanador(Sala)
        For indiceJugador = 0 To UBound(.Jugadores)
            If IsValidUserRef(.Jugadores(indiceJugador)) Then
                If UserList(.Jugadores(indiceJugador).ArrayIndex).flags.EquipoReto = ganador Then
                    Call WriteLocaleMsg(.Jugadores(indiceJugador).ArrayIndex, MSG_TIENES_MINUTO_LEVANTAR_ITEMS_PISO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1972=Tienes 1 minuto para levantar los items del piso.
                End If
            End If
        Next
        posBanquero.map = .PosIzquierda.map
        posBanquero.x = ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x
        posBanquero.y = ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y
        Call SetNpcRef(.Banquero, SpawnNpc(NPC_BANQUERO_RETO, posBanquero, True, False))
    End With
    Exit Sub
IniciarDepositoItems_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarDepositoItems", Erl)
End Sub

Public Sub TerminarTiempoAgarrarItems(ByVal Sala As Integer)
    On Error GoTo TerminarTiempoAgarrarItems_Err
    Dim indiceJugador As Integer
    Dim tUser         As t_UserReference
    Dim posX          As Byte
    Dim posY          As Byte
    With Retos.Salas(Sala)
        If .Fase <> e_FaseReto.DepositoItems Then Exit Sub
        .Fase = e_FaseReto.Finalizando
        .TiempoItems = 0
        If IsValidNpcRef(.Banquero) Then
            Call QuitarNPC(.Banquero.ArrayIndex, eChallenge)
        End If
        Call ClearNpcRef(.Banquero)
        ' Sacamos a los ganadores que siguen en la sala
        For indiceJugador = 0 To UBound(.Jugadores)
            tUser = .Jugadores(indiceJugador)
            If IsValidUserRef(tUser) Then
                If UserList(tUser.ArrayIndex).flags.EnReto And UserList(tUser.ArrayIndex).flags.SalaReto = Sala Then
                    Call RevivirYLimpiar(tUser.ArrayIndex)
                    Call SacarDeReto(tUser.ArrayIndex)
                    Call DevolverPosAnterior(tUser.ArrayIndex)
                End If
            End If
        Next
        For posX = .PosIzquierda.x To .PosDerecha.x
            For posY = .PosIzquierda.y To .PosDerecha.y
                Call EraseObj(GetMaxInvOBJ(), .PosIzquierda.map, posX, posY)
            Next
        Next
    End With
    Call SalaLiberada(Sala)
    Exit Sub
TerminarTiempoAgarrarItems_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.TerminarTiempoAgarrarItems", Erl)
End Sub

Public Sub AbandonarReto(ByVal UserIndex As Integer, Optional ByVal Desconexion As Boolean)
    On Error GoTo AbandonarReto_Err
    Dim Sala          As Integer
    Dim Equipo        As e_EquipoReto
    Dim indiceJugador As Integer
    Dim texto         As String
    Dim finalizar     As Boolean
    Sala = UserList(UserIndex).flags.SalaReto
    Equipo = UserList(UserIndex).flags.EquipoReto
    If Sala <= 0 Then
        Call SacarDeReto(UserIndex)
        Exit Sub
    End If
    ' En la fase de items el reto ya está pagado: solo sacamos al ganador
    If Retos.Salas(Sala).Fase = e_FaseReto.DepositoItems Then
        Call AbandonarDepositoItems(UserIndex, Sala)
        Exit Sub
    End If
    Call SacarDeReto(UserIndex)
    If Retos.Salas(Sala).Fase <> e_FaseReto.EnJuego Then Exit Sub
    If Not Desconexion Then
        Call WriteLocaleMsg(UserIndex, MSG_HAS_ABANDONADO_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1973=Has abandonado el reto.
    End If
    ' Restamos un miembro al equipo y si llega a cero entonces procesamos la derrota
    With Retos.Salas(Sala)
        If Equipo = e_EquipoReto.Izquierda Then
            If .TamañoEquipoIzq <= 1 Then
                .Puntaje = PUNTAJE_ABANDONO ' Forzamos la victoria del equipo derecho
                finalizar = True
            Else
                .TamañoEquipoIzq = .TamañoEquipoIzq - 1
            End If
        Else
            If .TamañoEquipoDer <= 1 Then
                .Puntaje = -PUNTAJE_ABANDONO ' Forzamos la victoria del equipo izquierdo
                finalizar = True
            Else
                .TamañoEquipoDer = .TamañoEquipoDer - 1
            End If
        End If
    End With
    If finalizar Then
        Call FinalizarReto(Sala)
        Exit Sub
    End If
    Call RevivirYLimpiar(UserIndex)
    Call DevolverPosAnterior(UserIndex)
    If Desconexion Then
        texto = UserList(UserIndex).name & " es descalificado por desconectarse."
    Else
        texto = UserList(UserIndex).name & " ha abandonado el reto."
    End If
    With Retos.Salas(Sala)
        For indiceJugador = 0 To UBound(.Jugadores)
            If .Jugadores(indiceJugador).ArrayIndex = UserIndex Then
                Call SetUserRef(.Jugadores(indiceJugador), 0)
            ElseIf IsValidUserRef(.Jugadores(indiceJugador)) Then
                Call WriteConsoleMsg(.Jugadores(indiceJugador).ArrayIndex, texto, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        Next
    End With
    ' Si en su equipo solo quedan muertos, la ronda es para el rival
    If Not EquipoTieneVivos(Sala, Equipo) Then
        Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    End If
    Exit Sub
AbandonarReto_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AbandonarReto", Erl)
End Sub

Private Sub SalaLiberada(ByVal Sala As Integer)
    On Error GoTo SalaLiberada_Err
    Dim Oferente       As Integer
    Dim NombreOferente As String
    With Retos.Salas(Sala)
        .EnUso = False
        .Fase = e_FaseReto.Inactiva
    End With
    Retos.SalasLibres = Retos.SalasLibres + 1
    ' Probamos la cola en orden hasta que alguno pueda arrancar en esta sala
    Do While ListaDeEspera.count > 0 And Not Retos.Salas(Sala).EnUso
        Oferente = ListaDeEspera.Keys(0)
        NombreOferente = ListaDeEspera.item(Oferente)
        Call ListaDeEspera.Remove(Oferente)
        If EsOferenteEnColaValido(Oferente, NombreOferente) Then
            UserList(Oferente).flags.SolicitudReto.Estado = e_SolicitudRetoEstado.Enviada
            Call IniciarReto(Oferente, Sala)
        End If
    Loop
    Exit Sub
SalaLiberada_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.SalaLiberada", Erl)
End Sub

Public Function PuedeReto(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then Exit Function
        If .flags.EnConsulta Then Exit Function
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then Exit Function
        If MapInfo(.pos.Map).Seguro = 0 Then Exit Function
        If .flags.EnTorneo Then Exit Function
        If MapData(.pos.x, .pos.y, .pos.Map).trigger = CARCEL Then Exit Function
    End With
    PuedeReto = True
End Function

Public Function PuedeRetoConMensaje(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .flags.EnReto Then
            Call WriteLocaleMsg(UserIndex, MSG_ENCUENTRAS_RETO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1974=Ya te encuentras en un reto.
            Exit Function
        End If
        If .flags.EnConsulta Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_ACCEDER_RETO_CONSULTA, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1975=No puedes acceder a un reto si estás en consulta.
            Exit Function
        End If
        If .flags.jugando_captura = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_JUGAR_RETO_ESTANDO_EVENTO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1976=No puedes jugar un reto estando en un evento.
            Exit Function
        End If
        If Not esCiudad(.pos.Map) Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_PARTICIPAR_RETO_MAPA_INSEGURO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1977=No puedes participar de un reto en un mapa inseguro.
            Exit Function
        End If
        If .flags.EnTorneo Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_PUEDES_IR_RETO_PARTICIPAS_TORNEO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1978=No puedes ir a un reto si participas de un torneo.
            Exit Function
        End If
        If MapData(.pos.x, .pos.y, .pos.Map).trigger = CARCEL Then
            Call WriteLocaleMsg(UserIndex, MSG_ENCARCELADO, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos, vbNullString) ' Msg1979=¡Estás encarcelado!
            Exit Function
        End If
    End With
    PuedeRetoConMensaje = True
End Function

Private Function IndiceJugadorEnSolicitud(ByVal UserIndex As Integer, ByVal Oferente As Integer) As Integer
    With UserList(Oferente).flags.SolicitudReto
        IndiceJugadorEnSolicitud = -1
        If .Estado <> e_SolicitudRetoEstado.Enviada Then Exit Function
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).nombre = UserList(UserIndex).name Then
                IndiceJugadorEnSolicitud = i
                Exit Function
            End If
        Next
    End With
End Function

Private Sub MensajeATodosSolicitud(ByVal Oferente As Integer, mensaje As String, ByVal Fuente As e_FontTypeNames)
    With UserList(Oferente).flags.SolicitudReto
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If .Jugadores(i).Aceptado Then
                Call WriteConsoleMsg(.Jugadores(i).CurIndex.ArrayIndex, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, Fuente)
            End If
        Next
        Call WriteConsoleMsg(Oferente, mensaje, e_TextChannel.TEXTCHANNEL_EVENT, Fuente)
    End With
End Sub

Private Function TodosPuedenReto(ByVal Oferente As Integer) As Boolean
    On Error GoTo ErrHandler
    With UserList(Oferente).flags.SolicitudReto
        If Not PuedeReto(Oferente) Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no puede entrar al reto en este momento.")
            Exit Function
        ElseIf UserList(Oferente).Stats.GLD < .Apuesta Then
            Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " no tiene las monedas de oro suficientes.")
            Exit Function
        ElseIf .PocionesMaximas >= 0 Then
            If TieneObjetos(OBJ_POCION_ROJA, .PocionesMaximas + 1, Oferente) Then
                Call CancelarSolicitudReto(Oferente, UserList(Oferente).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                Exit Function
            End If
        End If
        Dim i As Integer
        For i = 0 To UBound(.Jugadores)
            If Not PuedeReto(.Jugadores(i).CurIndex.ArrayIndex) Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no puede entrar al reto en este momento.")
                Exit Function
            ElseIf UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.GLD < .Apuesta Then
                Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no tiene las monedas de oro suficientes.")
                Exit Function
            ElseIf .PocionesMaximas >= 0 Then
                If TieneObjetos(OBJ_POCION_ROJA, .PocionesMaximas + 1, .Jugadores(i).CurIndex.ArrayIndex) Then
                    Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " tiene demasiadas pociones rojas (Cantidad máxima: " & _
                            .PocionesMaximas & ").")
                    Exit Function
                End If
            End If
        Next
        TodosPuedenReto = True
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModRetos.TodosPuedenReto", Erl)
End Function

Private Function EquipoContrario(ByVal Equipo As e_EquipoReto) As e_EquipoReto
    If Equipo = e_EquipoReto.Izquierda Then
        EquipoContrario = e_EquipoReto.Derecha
    Else
        EquipoContrario = e_EquipoReto.Izquierda
    End If
End Function

Private Function ObtenerTamañoEquipo(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As Integer
    If Equipo = e_EquipoReto.Izquierda Then
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoIzq
    Else
        ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoDer
    End If
End Function

Private Sub RevivirYLimpiar(ByVal UserIndex As Integer)
    Call WriteStopped(UserIndex, False)
    ' Si está vivo
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call LimpiarEstadosAlterados(UserIndex)
    End If
    ' Si está muerto lo revivimos, sino lo curamos
    Call RevivirUsuario(UserIndex)
End Sub


Private Sub SacarDeReto(ByVal UserIndex As Integer)
    On Error GoTo SacarDeReto_Err
    With UserList(UserIndex)
        .flags.EnReto = False
        .flags.SalaReto = 0
        .Counters.CuentaRegresiva = SIN_CUENTA_REGRESIVA
    End With
    Exit Sub
SacarDeReto_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.SacarDeReto", Erl)
End Sub

Private Function ObtenerGanador(ByVal Sala As Integer) As e_EquipoReto
    If Retos.Salas(Sala).Puntaje < 0 Then
        ObtenerGanador = e_EquipoReto.Izquierda
    Else
        ObtenerGanador = e_EquipoReto.Derecha
    End If
End Function

Private Function EquipoTieneVivos(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As Boolean
    On Error GoTo EquipoTieneVivos_Err
    Dim indiceJugador As Integer
    With Retos.Salas(Sala)
        ' El equipo izquierdo ocupa los índices pares y el derecho los impares
        For indiceJugador = IIf(Equipo = e_EquipoReto.Izquierda, 0, 1) To UBound(.Jugadores) Step 2
            If IsValidUserRef(.Jugadores(indiceJugador)) Then
                If UserList(.Jugadores(indiceJugador).ArrayIndex).flags.Muerto = 0 Then
                    EquipoTieneVivos = True
                    Exit Function
                End If
            End If
        Next
    End With
    Exit Function
EquipoTieneVivos_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.EquipoTieneVivos", Erl)
End Function

Private Function ArmarNombresEquipo(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As String
    On Error GoTo ArmarNombresEquipo_Err
    Dim indiceJugador As Integer
    Dim nombres       As String
    Dim ultimoNombre  As String
    With Retos.Salas(Sala)
        For indiceJugador = IIf(Equipo = e_EquipoReto.Izquierda, 0, 1) To UBound(.Jugadores) Step 2
            If IsValidUserRef(.Jugadores(indiceJugador)) Then
                If LenB(ultimoNombre) > 0 Then
                    nombres = nombres & IIf(LenB(nombres) > 0, ", ", vbNullString) & ultimoNombre
                End If
                ultimoNombre = UserList(.Jugadores(indiceJugador).ArrayIndex).name
            End If
        Next
    End With
    If LenB(nombres) > 0 Then
        ArmarNombresEquipo = nombres & " y " & ultimoNombre
    Else
        ArmarNombresEquipo = ultimoNombre
    End If
    Exit Function
ArmarNombresEquipo_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.ArmarNombresEquipo", Erl)
End Function

Private Sub AbandonarDepositoItems(ByVal UserIndex As Integer, ByVal Sala As Integer)
    On Error GoTo AbandonarDepositoItems_Err
    Dim indiceJugador   As Integer
    Dim tUser           As t_UserReference
    Dim quedanGanadores As Boolean
    Call SacarDeReto(UserIndex)
    Call RevivirYLimpiar(UserIndex)
    Call DevolverPosAnterior(UserIndex)
    With Retos.Salas(Sala)
        For indiceJugador = 0 To UBound(.Jugadores)
            tUser = .Jugadores(indiceJugador)
            If tUser.ArrayIndex = UserIndex Then
                Call SetUserRef(.Jugadores(indiceJugador), 0)
            ElseIf IsValidUserRef(tUser) Then
                If UserList(tUser.ArrayIndex).flags.EnReto And UserList(tUser.ArrayIndex).flags.SalaReto = Sala Then
                    quedanGanadores = True
                End If
            End If
        Next
    End With
    If Not quedanGanadores Then
        Call TerminarTiempoAgarrarItems(Sala)
    End If
    Exit Sub
AbandonarDepositoItems_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.AbandonarDepositoItems", Erl)
End Sub

Private Function EsOferenteEnColaValido(ByVal Oferente As Integer, ByVal NombreOferente As String) As Boolean
    On Error GoTo EsOferenteEnColaValido_Err
    Dim referencia As t_UserReference
    referencia = NameIndex(NombreOferente)
    If Not IsValidUserRef(referencia) Then Exit Function
    If referencia.ArrayIndex <> Oferente Then Exit Function
    If UserList(Oferente).flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.EnCola Then Exit Function
    EsOferenteEnColaValido = True
    Exit Function
EsOferenteEnColaValido_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.EsOferenteEnColaValido", Erl)
End Function

Private Function NombreRepetidoEnReto(ByVal NombreOferente As String, ByRef Jugadores() As String) As String
    On Error GoTo NombreRepetidoEnReto_Err
    Dim indiceJugador     As Integer
    Dim indiceComparacion As Integer
    For indiceJugador = 0 To UBound(Jugadores)
        If UCase$(Trim$(Jugadores(indiceJugador))) = UCase$(NombreOferente) Then
            NombreRepetidoEnReto = Jugadores(indiceJugador)
            Exit Function
        End If
        For indiceComparacion = indiceJugador + 1 To UBound(Jugadores)
            If UCase$(Trim$(Jugadores(indiceJugador))) = UCase$(Trim$(Jugadores(indiceComparacion))) Then
                NombreRepetidoEnReto = Jugadores(indiceJugador)
                Exit Function
            End If
        Next
    Next
    Exit Function
NombreRepetidoEnReto_Err:
    Call TraceError(Err.Number, Err.Description, "ModRetos.NombreRepetidoEnReto", Erl)
End Function
