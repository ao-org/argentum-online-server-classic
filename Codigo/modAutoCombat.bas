Attribute VB_Name = "modAutoCombat"
Option Explicit

Private Enum e_AutoRangedResult
    e_AutoRangedFired = 0
    e_AutoRangedWaiting = 1
    e_AutoRangedNoArrows = 2
    e_AutoRangedRejected = 3
End Enum

Public Sub ResetAutoCombat(ByVal UserIndex As Integer)
    On Error GoTo ResetAutoCombat_Err

    With UserList(UserIndex).AutoCombat
        .State = e_AutoCombatState.e_AutoCombatOff
        .actionType = e_AutoCombatActionType.e_AutoCombatActionNone
        .spellSlot = 0
        .SpellIndex = 0
        Call ClearNpcRef(.TargetNPC)
    End With

    Exit Sub

ResetAutoCombat_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.ResetAutoCombat", Erl)
End Sub

Public Function IsAutoCombatEnabled(ByVal UserIndex As Integer) As Boolean
    IsAutoCombatEnabled = (UserList(UserIndex).AutoCombat.State <> e_AutoCombatState.e_AutoCombatOff)
End Function

Public Sub StopAutoCombat(ByVal UserIndex As Integer, ByVal stopReason As e_AutoCombatStopReason)
    On Error GoTo StopAutoCombat_Err

    If Not IsAutoCombatEnabled(UserIndex) Then Exit Sub

    Call ResetAutoCombat(UserIndex)
    If stopReason <> e_AutoCombatStopReason.e_AutoCombatStopDisconnected Then
        Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatOff)
    End If
    Dim localeMsgId As Integer
    localeMsgId = GetAutoCombatStopMessage(stopReason)

    If localeMsgId > 0 Then
        Call WriteLocaleMsg(UserIndex, localeMsgId, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub

StopAutoCombat_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.StopAutoCombat", Erl)
End Sub

Private Function GetAutoCombatStopMessage(ByVal stopReason As e_AutoCombatStopReason) As Integer
    Select Case stopReason
        Case e_AutoCombatStopReason.e_AutoCombatStopByUser
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_OFF
        Case e_AutoCombatStopReason.e_AutoCombatStopTargetDied
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_TARGET_DIED
        Case e_AutoCombatStopReason.e_AutoCombatStopOutOfRange
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_OUT_OF_RANGE
        Case e_AutoCombatStopReason.e_AutoCombatStopOtherTarget
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_OTHER_TARGET
        Case e_AutoCombatStopReason.e_AutoCombatStopNoArrows
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_NO_ARROWS
        Case e_AutoCombatStopReason.e_AutoCombatStopDamageReceived
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_DAMAGE_RECEIVED
        Case e_AutoCombatStopReason.e_AutoCombatStopMeditationCut
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_MEDITATION_CUT
        Case e_AutoCombatStopReason.e_AutoCombatStopCantMeditate
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_CANT_MEDITATE
        Case e_AutoCombatStopReason.e_AutoCombatStopUserDied, _
             e_AutoCombatStopReason.e_AutoCombatStopMapChanged
            GetAutoCombatStopMessage = MSG_AUTOCOMBAT_OFF
        Case Else
            GetAutoCombatStopMessage = 0
    End Select
End Function

Public Sub ToggleAutoCombat(ByVal UserIndex As Integer)
    On Error GoTo ToggleAutoCombat_Err

    If IsAutoCombatEnabled(UserIndex) Then
        Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopByUser)
        Exit Sub
    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    Call ResetAutoCombat(UserIndex)
    UserList(UserIndex).AutoCombat.State = e_AutoCombatState.e_AutoCombatArmed

    Call WriteLocaleMsg(UserIndex, MSG_AUTOCOMBAT_ARMED, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_INFO)
    Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatArmed)
    Exit Sub

ToggleAutoCombat_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.ToggleAutoCombat", Erl)
End Sub

Public Sub OnUserCombatActionOnNpc(ByVal UserIndex As Integer, ByVal actionType As e_AutoCombatActionType, ByVal npcIndex As Integer, Optional ByVal spellSlot As Byte = 0)
    On Error GoTo OnUserCombatActionOnNpc_Err

    If npcIndex <= 0 Then Exit Sub

    With UserList(UserIndex).AutoCombat
        Select Case .State
            Case e_AutoCombatState.e_AutoCombatOff
                Exit Sub

            Case e_AutoCombatState.e_AutoCombatArmed
                .actionType = actionType
                .spellSlot = spellSlot
                If spellSlot > 0 Then .SpellIndex = UserList(UserIndex).Stats.UserHechizos(spellSlot)
                Call SetNpcRef(.TargetNPC, npcIndex)
                .State = e_AutoCombatState.e_AutoCombatRepeating

                Call WriteLocaleMsg(UserIndex, MSG_AUTOCOMBAT_CAPTURED, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_INFO, NpcList(npcIndex).name)
                Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatRepeating)

            Case Else
                ' Acciones manuales sobre el mismo NPC se ignoran; sobre otro, se apaga.
                If .TargetNPC.ArrayIndex <> npcIndex Or Not IsValidNpcRef(.TargetNPC) Then
                    Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopOtherTarget)
                End If
        End Select
    End With

    Exit Sub

OnUserCombatActionOnNpc_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.OnUserCombatActionOnNpc", Erl)
End Sub

Public Sub OnUserMeditationInterrupted(ByVal UserIndex As Integer)
    On Error GoTo OnUserMeditationInterrupted_Err

    If UserList(UserIndex).AutoCombat.State <> e_AutoCombatState.e_AutoCombatMeditating Then Exit Sub
    Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopMeditationCut)
    Exit Sub

OnUserMeditationInterrupted_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.OnUserMeditationInterrupted", Erl)
End Sub

Private Function GetNpcInFrontOfUser(ByVal UserIndex As Integer) As Integer
    Dim frontPos As t_WorldPos
    frontPos = UserList(UserIndex).pos
    Call HeadtoPos(UserList(UserIndex).Char.Heading, frontPos)
    If Not InMapBounds(frontPos.map, frontPos.x, frontPos.y) Then Exit Function
    GetNpcInFrontOfUser = MapData(frontPos.x, frontPos.y, frontPos.map).npcIndex
End Function

Public Function GetMeleeTargetNpc(ByVal UserIndex As Integer) As Integer
    GetMeleeTargetNpc = GetNpcInFrontOfUser(UserIndex)
End Function

Public Sub OnAutoCombatMeditationCompleted(ByVal UserIndex As Integer)
    On Error GoTo OnAutoCombatMeditationCompleted_Err

    With UserList(UserIndex).AutoCombat
        If .State <> e_AutoCombatState.e_AutoCombatMeditating Then Exit Sub
        .State = e_AutoCombatState.e_AutoCombatRepeating
    End With

    Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatRepeating)
    Exit Sub

OnAutoCombatMeditationCompleted_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.OnAutoCombatMeditationCompleted", Erl)
End Sub



' Duplicado de HandleWorkLeftClick / Case e_Skill.Proyectiles (rama NPC).
' Cualquier cambio en el disparo manual debe replicarse aca.
Private Function TryAutoRangedAttack(ByVal UserIndex As Integer, ByVal npcIndex As Integer) As e_AutoRangedResult
    On Error GoTo TryAutoRangedAttack_Err

    Dim weaponInfo        As t_ObjData
    Dim projectileType    As Byte
    Dim invalidAmmoReason As Integer
    Dim munitionSlot      As Integer
    Dim consumeAmmo       As Boolean
    Dim TargetX           As Byte
    Dim TargetY           As Byte

    TryAutoRangedAttack = e_AutoRangedWaiting

    With UserList(UserIndex)
        If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Function
        If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Function

        With .invent
            If .EquippedWeaponObjIndex < 1 Then
                TryAutoRangedAttack = e_AutoRangedNoArrows
                Exit Function
            End If
            weaponInfo = ObjData(.EquippedWeaponObjIndex)
            If IsItemInCooldown(UserList(UserIndex), .Object(.EquippedWeaponSlot)) Then Exit Function

            If weaponInfo.Proyectil = 1 And weaponInfo.Municion = 0 Then
                invalidAmmoReason = 0
            ElseIf .EquippedMunitionSlot < 1 Or .EquippedMunitionSlot > UserList(UserIndex).CurrentInventorySlots Then
                invalidAmmoReason = 1
            ElseIf .EquippedMunitionObjIndex = 0 Then
                invalidAmmoReason = 1
            ElseIf weaponInfo.Proyectil <> 1 Then
                invalidAmmoReason = 2
            ElseIf ObjData(.EquippedMunitionObjIndex).OBJType <> e_OBJType.otArrows Then
                invalidAmmoReason = 1
            ElseIf .Object(.EquippedMunitionSlot).Amount < 1 Then
                invalidAmmoReason = 1
            ElseIf ObjData(.EquippedMunitionObjIndex).Subtipo <> weaponInfo.Municion Then
                invalidAmmoReason = 1
            End If

            If invalidAmmoReason <> 0 Then
                If invalidAmmoReason = 1 Then
                    Call WriteLocaleMsg(UserIndex, MSG_NO_TENES_MUNICIONES, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
                    Call Desequipar(UserIndex, .EquippedMunitionSlot)
                End If
                TryAutoRangedAttack = e_AutoRangedNoArrows
                Exit Function
            End If
        End With

        ' Sin stamina: espera sin mensaje para no spamear la consola cada tick.
        If .Stats.MinSta < 10 Then Exit Function

        ' El intervalo del arco se consume recien aca, cuando el disparo es seguro.
        If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Function

        Call QuitarSta(UserIndex, RandomNumber(1, 10))

        If NpcList(npcIndex).Attackable = 0 Then
            TryAutoRangedAttack = e_AutoRangedRejected
            Exit Function
        End If

        Dim AttackResult As t_AttackInteractionResult
        AttackResult = UserCanAttackNpc(UserIndex, npcIndex)
        Call SendAttackInteractionMessage(UserIndex, AttackResult.result)
        If Not AttackResult.CanAttack Then
            TryAutoRangedAttack = e_AutoRangedRejected
            Exit Function
        End If

        If AttackResult.TurnPK Then Call VolverCriminal(UserIndex)

        projectileType = GetProjectileView(UserList(UserIndex))
        TargetX = NpcList(npcIndex).pos.x
        TargetY = NpcList(npcIndex).pos.y

        Call UsuarioAtacaNpc(UserIndex, npcIndex, Ranged)
        consumeAmmo = True

        If projectileType > 0 And .flags.Oculto = 0 Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(.pos.x, .pos.y, TargetX, TargetY, projectileType))
        End If
        If .flags.AdminInvisible = 0 Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex, 1))
        End If

        With .invent
            If weaponInfo.Proyectil = 1 And weaponInfo.Municion > 0 Then
                munitionSlot = .EquippedMunitionSlot
                If ObjData(.EquippedWeaponObjIndex).CreaWav > 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(ObjData(.EquippedWeaponObjIndex).CreaWav, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                End If
                If munitionSlot <> 0 Then
                    If consumeAmmo And Not IsConsumableFreeZone(UserIndex) Then
                        Call QuitarUserInvItem(UserIndex, munitionSlot, 1)
                    End If
                    If .Object(munitionSlot).Amount > 0 Then
                        .EquippedMunitionSlot = munitionSlot
                        .EquippedMunitionObjIndex = .Object(munitionSlot).ObjIndex
                        .Object(munitionSlot).Equipped = 1
                    Else
                        .EquippedMunitionSlot = 0
                        .EquippedMunitionObjIndex = 0
                    End If
                    Call UpdateUserInv(False, UserIndex, munitionSlot)
                End If
            ElseIf consumeAmmo Then
                Call UpdateCd(UserIndex, weaponInfo.cdType)
            End If
        End With
    End With

    TryAutoRangedAttack = e_AutoRangedFired
    Exit Function

TryAutoRangedAttack_Err:
    TryAutoRangedAttack = e_AutoRangedRejected
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.TryAutoRangedAttack", Erl)
End Function

Private Sub ProcessAutoMelee(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
    On Error GoTo ProcessAutoMelee_Err

    ' Solo pega con el NPC adyacente y de frente; si no, espera.
    If GetMeleeTargetNpc(UserIndex) <> npcIndex Then Exit Sub
    If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
    If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

    Call UsuarioAtaca(UserIndex)
    Exit Sub

ProcessAutoMelee_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.ProcessAutoMelee", Erl)
End Sub

Public Sub ProcessAutoCombat(ByVal UserIndex As Integer)
    On Error GoTo ProcessAutoCombat_Err

    Dim npcIndex As Integer

    With UserList(UserIndex).AutoCombat
        If .State = e_AutoCombatState.e_AutoCombatOff Then Exit Sub
        If .State = e_AutoCombatState.e_AutoCombatArmed Then Exit Sub

        If Not IsValidNpcRef(.TargetNPC) Then
            Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopTargetDied)
            Exit Sub
        End If

        npcIndex = .TargetNPC.ArrayIndex

        If Not NpcList(npcIndex).flags.NPCActive Or NpcList(npcIndex).Stats.MinHp <= 0 Then
            Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopTargetDied)
            Exit Sub
        End If

        If Not IsAutoCombatTargetInVision(UserIndex, npcIndex) Then
            Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopOutOfRange)
            Exit Sub
        End If

        If .State = e_AutoCombatState.e_AutoCombatMeditating Then
            ' Si la meditacion se corto por una via sin hook, se trata como corte.
            If Not UserList(UserIndex).flags.Meditando Then
                Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopMeditationCut)
            End If
            Exit Sub
        End If

        ' Meditando por cualquier via: no actuar hasta que termine.
        If UserList(UserIndex).flags.Meditando Then
            .State = e_AutoCombatState.e_AutoCombatMeditating
            Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatMeditating)
            Exit Sub
        End If

        Select Case .actionType
            Case e_AutoCombatActionType.e_AutoCombatActionMelee
                Call ProcessAutoMelee(UserIndex, npcIndex)

            Case e_AutoCombatActionType.e_AutoCombatActionRanged
                Select Case TryAutoRangedAttack(UserIndex, npcIndex)
                    Case e_AutoRangedResult.e_AutoRangedNoArrows
                        Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopNoArrows)
                    Case e_AutoRangedResult.e_AutoRangedRejected
                        Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopByUser)
                End Select

            Case e_AutoCombatActionType.e_AutoCombatActionSpell
                Call ProcessAutoSpell(UserIndex, npcIndex)
        End Select
    End With

    Exit Sub

ProcessAutoCombat_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.ProcessAutoCombat", Erl)
End Sub

Private Function IsAutoCombatTargetInVision(ByVal UserIndex As Integer, ByVal npcIndex As Integer) As Boolean
    With NpcList(npcIndex).pos
        If .map <> UserList(UserIndex).pos.map Then Exit Function
        If Abs(.x - UserList(UserIndex).pos.x) > RANGO_VISION_X Then Exit Function
        If Abs(.y - UserList(UserIndex).pos.y) > RANGO_VISION_Y Then Exit Function
    End With
    IsAutoCombatTargetInVision = True
End Function

Private Sub ProcessAutoSpell(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
    On Error GoTo ProcessAutoSpell_Err

    Dim castSpellSlot  As Byte
    Dim castSpellIndex As Integer

    With UserList(UserIndex)
        castSpellSlot = .AutoCombat.spellSlot
        castSpellIndex = .AutoCombat.SpellIndex

        ' El hechizo del slot cambio o fue reordenado.
        If .Stats.UserHechizos(castSpellSlot) <> castSpellIndex Then
            Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopByUser)
            Exit Sub
        End If

        ' Sin mana: meditacion automatica.
        If .Stats.MinMAN < Hechizos(castSpellIndex).ManaRequerido Then
            If Not TryStartMeditation(UserIndex) Then
                Call StopAutoCombat(UserIndex, e_AutoCombatStopReason.e_AutoCombatStopCantMeditate)
                Exit Sub
            End If
            .AutoCombat.State = e_AutoCombatState.e_AutoCombatMeditating
            Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatMeditating)
            Exit Sub
        End If

        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub

        ' Mismo patron que ExecuteHooTargetedSpellCast: fijar objetivo, lanzar y restaurar.
        Dim previousSpell      As Integer
        Dim previousTargetNpc  As t_NpcReference
        Dim previousTargetUser As t_UserReference

        previousSpell = .flags.Hechizo
        previousTargetNpc = .flags.TargetNPC
        previousTargetUser = .flags.targetUser

        .flags.Hechizo = castSpellSlot
        Call ClearUserRef(.flags.targetUser)
        Call SetNpcRef(.flags.TargetNPC, npcIndex)

        .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
        Call LanzarHechizo(castSpellSlot, UserIndex)

        .flags.Hechizo = previousSpell
        Call ClearNpcRef(.flags.TargetNPC)
        Call ClearUserRef(.flags.targetUser)
        If IsValidNpcRef(previousTargetNpc) Then Call SetNpcRef(.flags.TargetNPC, previousTargetNpc.ArrayIndex)
        If IsValidUserRef(previousTargetUser) Then Call SetUserRef(.flags.targetUser, previousTargetUser.ArrayIndex)
    End With

    Exit Sub

ProcessAutoSpell_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.ProcessAutoSpell", Erl)
End Sub

Public Sub OnUserMeditationStarted(ByVal UserIndex As Integer)
    On Error GoTo OnUserMeditationStarted_Err

    With UserList(UserIndex).AutoCombat
        If .State <> e_AutoCombatState.e_AutoCombatRepeating Then Exit Sub
        .State = e_AutoCombatState.e_AutoCombatMeditating
    End With

    Call WriteAutoCombatState(UserIndex, e_AutoCombatState.e_AutoCombatMeditating)
    Exit Sub

OnUserMeditationStarted_Err:
    Call TraceError(Err.Number, Err.Description, "modAutoCombat.OnUserMeditationStarted", Erl)
End Sub
