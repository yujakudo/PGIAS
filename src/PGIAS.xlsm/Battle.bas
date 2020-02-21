Attribute VB_Name = "Battle"

Option Explicit

'   ������ʂ̃C���f�b�N�X
Const IDX_SelfAtk As Integer = 0
Const IDX_SelfDef As Integer = 1
Const IDX_EnemyAtk As Integer = 2
Const IDX_EnemyDef As Integer = 3

Dim InterTypeFactorTable As Variant
Dim WeatherFactorTable As Variant
Dim TypeMatchFactorValue As Double
Dim MatchBattleFactorValue As Double
Dim ChargeByDamage As Double

'   �ʏ�킴
Public Type NormalAttack
    idleTime As Double
    charge As Integer
End Type

'   ����
Public Type AttackEffect
    desc As String
    probability As Double
    step As Integer
    factor As Variant
    stages As Variant
    expect As Variant
End Type

'   �Q�[�W�킴
Public Type SpecialAttack
    rsvNum As Integer
    rsvVol As Double
    curVol As Double
End Type

'   �킴
Public Type Attack
    name As String
    itype As Integer
    power As Double
    damage As Double
    idleTime As Double
    class As Integer
    normal As NormalAttack
    special As SpecialAttack
    isEffect As Boolean
    effect As AttackEffect
    flag As Integer
End Type

'   �킴�̃C���f�b�N�X
Public Type AttackIndex
    selected As Integer
    lower As Integer
    upper As Integer
End Type

'   ����ɗ^�����_���[�W�̗ݐ�
Public Type GivenDamage
    time As Double
    damage As Double
End Type

'   �|�P����
Public Type Monster
    nickname As String
    species As String
    logName As String
    itype(1) As Integer
    PL As Double
    indATK As Integer
    indDEF As Integer
    indHP As Integer
    atkPower As Double
    defPower As Double
    hpPower As Double
    fullHP As Long
    curHP As Double
    CP As Integer
    attacks() As Attack
    atkIndex(1) As AttackIndex
    given(1) As GivenDamage
    atkNum As Integer
    chargeCount As Double
    clock As Double
    mode As Integer
    phase As Integer
End Type

'   �Z�̕]���l
Public Type AttackParam
    name As String
    class As Integer
    damage As Double
    damageEfc As Double
    chargeEfc As Double
End Type

Public Type CDpsSet
    natk As String
    satk As String
    cDps As Double
    Cycle As Double
End Type



'   �_���[�W�̌v�Z
Public Function getDamage(ByVal mode As String, _
        ByVal species As String, ByVal atkType As String, _
        ByVal atkName As String, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal PL As Double = 40, _
        Optional ByVal weather As String = "", _
        Optional ByVal enemySpecies As String = "", _
        Optional ByVal enemyPL As Double = 40, _
        Optional ByVal enemyIndDef As Long = 15) As Long
    getDamage = Fix(getAnaDamage(mode, species, atkType, atkName, _
        indATK, PL, weather, enemySpecies, enemyPL, enemyIndDef))
End Function

'   �_���[�W�̌v�Z�i��͗p�j
Public Function getAnaDamage(ByVal mode As Integer, _
        ByVal species As String, ByVal atkClass As Variant, _
        ByVal atkName As String, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal PL As Double = 40, _
        Optional ByVal weather As String = "", _
        Optional ByVal enemySpecies As String = "", _
        Optional ByVal enemyPL As Double = 40, _
        Optional ByVal enemyIndDef As Long = 15) As Double
    Dim self As Monster
    Dim enemy As Monster
    Dim natkName, satkName As String
    Dim atkIdx As Integer
    
    If Not IsNumeric(atkClass) Then atkClass = getAtkClassIndex(atkClass)
    If Not IsNumeric(weather) Then weather = getWeatherIndex(weather)
    If species = "" Then Exit Function
    If atkClass = C_IdNormalAtk Then
        natkName = atkName
    Else
        satkName = atkName
    End If
    Call getMonster(self, species, PL, indATK)
    Call setAttacks(mode, self, natkName, satkName)
    Call getMonster(enemy, enemySpecies, enemyPL, 0, enemyIndDef)
    atkIdx = getAttackIndex(self, atkClass)
    getAnaDamage = calcADamage(atkIdx, self, enemy, True, weather)
End Function

'   �I������Ă���Z�̃C���f�b�N�X�̎擾
Private Function getAttackIndex(ByRef mon As Monster, _
                                Optional ByVal atkClass As Integer = -1) As Variant
    If atkClass = C_IdNormalAtk Then
        getAttackIndex = mon.atkIndex(0).selected
    ElseIf atkClass = C_IdSpecialAtk Then
        getAttackIndex = mon.atkIndex(1).selected
    Else
        getAttackIndex = Array(mon.atkIndex(0).selected, mon.atkIndex(1).selected)
    End If
End Function

'   cDPS�̌v�Z
Public Function getCDPS(ByVal mode As Integer, _
        ByVal species As String, _
        ByVal natkName As String, ByVal satkName As String, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal PL As Double = 40, _
        Optional ByVal weather As String = "", _
        Optional ByVal enemySpecies As String = "", _
        Optional ByVal enemyPL As Double = 40, _
        Optional ByVal enemyIndDef As Long = 15, _
        Optional ByVal isAna As Boolean = False) As Double
    Dim self As Monster
    Dim enemy As Monster
    Dim cDpss As CDpsSet
    Call getMonster(self, species, PL, indATK)
    Call setAttacks(mode, self, natkName, satkName)
    Call getMonster(enemy, enemySpecies, enemyPL, 0, enemyIndDef)
    cDpss = calcCDPS(self, enemy, isAna, weather)
    getCDPS = cDpss.cDps
End Function

'   cDPS�̎Z�o
Public Function calcCDPS(ByRef self As Monster, ByRef enemy As Monster, _
                Optional ByVal isAna As Boolean = False, _
                Optional ByVal weather As Variant = 0) As CDpsSet
    Dim damage, period As Double
    Dim atkIdx As Variant
    
    atkIdx = getAttackIndex(self)
    With self
        calcCDPS.natk = .attacks(atkIdx(0)).name
        calcCDPS.satk = .attacks(atkIdx(1)).name
    End With
    If Not IsNumeric(weather) Then weather = getWeatherIndex(weather)
    Call calcDamages(self, enemy, isAna, weather)
    If calcChargeCount(self) Then
        With self
            damage = .attacks(atkIdx(0)).damage * self.chargeCount + .attacks(atkIdx(1)).damage
            period = .attacks(atkIdx(0)).idleTime * self.chargeCount + .attacks(atkIdx(1)).idleTime
        End With
        If period = 0 Then
            Debug.Print "Period is 0 in calcCDPS. " & self.nickname
        End If
        calcCDPS.cDps = damage / period
        calcCDPS.Cycle = period
    End If
End Function

'   �U����̓p�����[�^�̎擾
'   �߂�l�̔z��̗v�f��1���珇�ɁA�_���[�W�A�_���[�W�����A�`���[�W�����i�ʏ�킴�̂݁j
Public Function getAtkAna( _
        ByRef atkParam() As AttackParam, _
        ByRef cDpss() As CDpsSet, _
        ByVal mode As Integer, _
        ByVal species As String, _
        Optional ByVal withLimited As Boolean = True, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal PL As Double = 40, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal enemySpecies As String = "", _
        Optional ByVal enemyPL As Double = 40, _
        Optional ByVal enemyIndDef As Long = 15)
    Dim atkNames, cDps As Variant
    Dim self As Monster
    Dim enemy As Monster
    Dim i, j, atkClass, ni, si As Integer
    Dim tmp As CDpsSet
    
    Call getMonster(self, species, PL, indATK)
    Call getMonster(enemy, enemySpecies, enemyPL, 0, enemyIndDef)
    '   ���ׂĂ̋Z��ݒ�
    atkNames = getAtkNames(species, False, withLimited)
    Call setAttacks(mode, self, atkNames(0), atkNames(1))
    ReDim atkParam(self.atkNum - 1)
    '   �N���X���ƁA���ׂĂ̋Z�̃p�����[�^���v�Z
    For i = 0 To self.atkNum - 1
        atkParam(i) = getAtkParam(self, enemy, i, weather)
    Next
    '   ���ׂĂ̋Z�g�ݍ��킹�ɂ���cDSP���v�Z�B�\�[�g�����Ă���
    With self
        ReDim cDpss((.atkIndex(0).upper - .atkIndex(0).lower + 1) _
                    * (.atkIndex(1).upper - .atkIndex(1).lower + 1) - 1)
        i = 0
        For ni = .atkIndex(0).lower To .atkIndex(0).upper
            self.atkIndex(0).selected = ni
            For si = .atkIndex(1).lower To .atkIndex(1).upper
                self.atkIndex(1).selected = si
                cDpss(i) = calcCDPS(self, enemy, True, weather)
                For j = i - 1 To 0 Step -1
                    If cDpss(j).cDps < cDpss(j + 1).cDps Then
                        tmp = cDpss(j): cDpss(j) = cDpss(j + 1): cDpss(j + 1) = tmp
                    End If
                Next
                i = i + 1
            Next
        Next
    End With
End Function


'   �U����̓p�����[�^�̎擾�i�_���[�W�A�_���[�W�����A�`���[�W�����j
Private Function getAtkParam(ByRef self As Monster, ByRef enemy As Monster, _
        ByVal atkIdx As Integer, _
        Optional ByVal weather As Integer = 0) As AttackParam
    Dim atkClass As Integer
    Dim ret() As Double
    Dim attrs() As String
    Dim vals As Variant
    
    With self.attacks(atkIdx)
        atkClass = .class
        '   �K�v�ȑ����̎擾
        If atkClass = C_IdNormalAtk Then   '   �ʏ�킴
            ReDim attrs(1), ret(2)
            '   �_���[�W�����̕���ƃ`���[�W����
            If self.mode = C_IdGym Then    '   ��������, �`���[�W���xEPS
                attrs(0) = ATK_IdleTime: attrs(1) = ATK_EPS
            ElseIf self.mode = C_IdMtc Then    '   �����^�[��, �`���[�W���xEPT
                attrs(0) = ATK_IdleTurnNum: attrs(1) = ATK_EPT
            End If
        Else    '   �Q�[�W�킴
            ReDim attrs(0), ret(1)
            '   �_���[�W�����̕���
            If self.mode = C_IdGym Then
                attrs(0) = ATK_IdleTime
            ElseIf self.mode = C_IdMtc Then
                attrs(0) = ATK_GaugeVolume
            End If
        End If
        vals = getAtkAttrs(atkClass, .name, attrs)
        getAtkParam.name = .name
        getAtkParam.class = atkClass
    End With
    '   Damage
    getAtkParam.damage = calcADamage(atkIdx, self, enemy, True, weather)
    '   �_���[�W����
    getAtkParam.damageEfc = getAtkParam.damage / vals(0)
    '   �`���[�W����
    If atkClass = C_IdNormalAtk Then getAtkParam.chargeEfc = vals(1)
End Function

'   �퓬�V�~�����[�V����
'   �Ԃ�l�̔z��̗v�f�́A0:KTR, 1:��KT, 2:�GKT, 3:���O
Private Function simBattle(ByRef sbjMon As Monster, ByRef objMon As Monster, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal needKTR As Boolean = False, _
        Optional ByVal needLog As Boolean = False) As Variant
    Dim sbjKT, objKT, DPS(2), rsv As Double
    Dim alive(1), isSbjFirst As Boolean
    Dim log, subLog As String
    
    sbjMon.clock = 0
    objMon.clock = 0
    sbjKT = 0: objKT = 0
    alive(0) = True: alive(1) = True
    Call calcDamages(sbjMon, objMon, False, weather)
    Call calcDamages(objMon, sbjMon, False, weather)
    isSbjFirst = sbjMon.atkPower >= objMon.atkPower
    '   �U�������肵�Ă���
    Call decideAttack(sbjMon)
    Call decideAttack(objMon)
    '   �ǂ��炩�������Ă����
    Do While alive(0) Or alive(1)
        If needLog Then
            log = log & subLog _
                & monStatusStr(sbjMon) & "  " & monStatusStr(objMon) & vbCrLf
        End If
        If sbjMon.clock < objMon.clock Or (isSbjFirst And sbjMon.clock = objMon.clock) Then
            subLog = hitAttack(sbjMon, objMon, weather) & vbCrLf
            '   �����HP��0�ȉ��ɂȂ�����KT���L�^
            If alive(1) And objMon.curHP <= 0 Then
                sbjKT = sbjMon.clock
                alive(1) = False
                If Not needKTR Then Exit Do
            End If
            Call decideAttack(sbjMon)
        Else
            subLog = hitAttack(objMon, sbjMon, weather) & vbCrLf
            If alive(0) And sbjMon.curHP <= 0 Then
                objKT = objMon.clock
                alive(0) = False
                If Not needKTR Then Exit Do
            End If
            Call decideAttack(objMon)
        End If
    Loop
    If needLog Then
        log = log & subLog _
            & monStatusStr(sbjMon) & "  " & monStatusStr(objMon) & vbCrLf
    End If
    If needKTR Then
        simBattle = Array(sbjKT / objKT, sbjKT, objKT, log)
    Else
        simBattle = Array(0, sbjKT, objKT, log)
    End If
End Function

'   ��Ԃ̕�����̍쐬
Private Function monStatusStr(ByRef mon As Monster) As String
    Dim rsv As Double
    With mon
        rsv = .attacks(.atkIndex(1).selected).special.curVol
        monStatusStr = msgstr(msgMonStatus, Array(.logName, .curHP, rsv))
    End With
End Function

'   ���̍U���̌���
Private Sub decideAttack(ByRef offence As Monster)
    Dim idleTime As Double
    Dim atkIdx As Integer
    
    atkIdx = offence.atkIndex(1).selected
    With offence.attacks(atkIdx)    '   �Q�[�W�Z
        '   �Q�[�W�������Ă��Ȃ���΁A�ʏ�Z���J�n
        If .special.curVol < .special.rsvVol Then atkIdx = offence.atkIndex(0).selected
    End With
    '   �U�������t�F�C�Y
    With offence.attacks(atkIdx)
        offence.phase = .class + 1
        idleTime = .idleTime
        '   �N���b�N
        offence.clock = offence.clock + idleTime
        With offence.given(.class)
            .time = .time + idleTime
        End With
    End With
End Sub

'   �U��
Private Function hitAttack(ByRef offence As Monster, ByRef deffence As Monster, _
        Optional ByVal weather As Integer = 0) As String
    Dim charge, damage As Double
    Dim name As String
    Dim atkIdx As Integer
    
    atkIdx = offence.atkIndex(offence.phase - 1).selected
    With offence.attacks(atkIdx)
        '   �_���[�W
        damage = .damage
        deffence.curHP = deffence.curHP - damage
        With offence.given(.class)
            .damage = .damage + damage
        End With
        '   �_���[�W�ɂ��h�䑤�̃`���[�W
        With deffence.attacks(deffence.atkIndex(1).selected)
            .special.curVol = .special.curVol + damage * ChargeByDamage
        End With
        If .class = 0 Then  '   �ʏ�Z
            '   �`���[�W
            charge = .normal.charge
            With offence.attacks(offence.atkIndex(1).selected)
                .special.curVol = .special.curVol + charge
            End With
        Else  '   �Q�[�W�Z
            '   �W����ŃQ�[�W��1�Ȃ�Q�[�W�����
            If offence.mode = C_IdGym And .special.rsvNum = 1 Then
                .special.curVol = 0
            Else
                .special.curVol = .special.curVol - .special.rsvVol
            End If
            '   ������ʂ�����΃_���[�W���Čv�Z
            If .isEffect Then
                Call stochTrans(.effect)
                Call calcDamages(offence, deffence, False, weather)
                Call calcDamages(deffence, offence, False, weather)
            End If
        End If
        hitAttack = msgstr(msgHitAttack, Array(offence.logName, .name, damage))
    End With
    offence.phase = 0
End Function


'   �ʏ�킴�̐ݒ�B�idamage�̓Z�b�g���Ȃ��j
Public Sub setNormalAttacks(ByRef mon As Monster, _
                    ByRef atkNames As Variant, _
                    Optional ByVal atkDelay As Double = 0, _
                    Optional ByVal flags As Variant = Nothing)
    Dim nattr As Object
    Dim i, idx, atkNum As Integer
    
    atkNum = UBound(atkNames) + 1
    ReDim Preserve mon.attacks(mon.atkNum + atkNum - 1)
    idx = mon.atkNum
    For i = 0 To atkNum - 1
        If atkNames(i) <> "" Then
            With mon.attacks(idx)
                .name = atkNames(i)
                .class = C_IdNormalAtk
                If IsArray(flags) Then .flag = flags(i)
                Set nattr = getAtkAttrs(C_NormalAttack, .name)
                .itype = getTypeIndex(nattr(ATK_Type))
                If mon.mode = C_IdGym Then
                    .power = nattr(ATK_GymPower)
                    .idleTime = nattr(ATK_IdleTime) + atkDelay
                    .normal.charge = nattr(ATK_GymCharge)
                ElseIf mon.mode = C_IdMtc Then
                    .power = nattr(ATK_MtcPower)
                    .idleTime = nattr(ATK_IdleTurnNum) + atkDelay
                    .normal.charge = nattr(ATK_MtcCharge)
                End If
            End With
            idx = idx + 1
        End If
    Next
    Call setAtkIndexes(mon, 0, idx)
End Sub

'   �Q�[�W�킴�̐ݒ�B�idamage�̓Z�b�g���Ȃ��j
Public Sub setSpecialAttacks(ByRef mon As Monster, _
                    ByRef atkNames As Variant, _
                    Optional ByVal atkDelay As Double = 0, _
                    Optional ByVal considerEffect As Boolean = False, _
                    Optional ByVal flags As Variant = Nothing)
    Dim sattr As Object
    Dim i, idx, atkNum As Integer
    
    atkNum = UBound(atkNames) + 1
    ReDim Preserve mon.attacks(mon.atkNum + atkNum - 1)
    idx = mon.atkNum
    For i = 0 To atkNum - 1
        If atkNames(i) <> "" Then
            With mon.attacks(idx)
                .name = atkNames(i)
                .class = C_IdSpecialAtk
                If IsArray(flags) Then .flag = flags(i)
                Set sattr = getAtkAttrs(C_SpecialAttack, .name)
                .itype = getTypeIndex(sattr(ATK_Type))
                If mon.mode = C_IdGym Then
                    .power = sattr(ATK_GymPower)
                    .idleTime = sattr(ATK_IdleTime) + atkDelay
                    .special.rsvNum = sattr(ATK_GaugeNumber)
                    If .special.rsvNum = 0 Then
                        Debug.Print "Number of Reserver is 0. " & atkNames(i)
                    End If
                    .special.rsvVol = 100# / .special.rsvNum
                ElseIf mon.mode = C_IdMtc Then
                    .power = sattr(ATK_MtcPower)
                    .idleTime = 1 + atkDelay
                    .special.rsvNum = 1
                    .special.rsvVol = sattr(ATK_GaugeVolume)
                End If
                .isEffect = initEffectTrans(.effect, sattr, considerEffect)
            End With
            idx = idx + 1
        End If
    Next
    Call setAtkIndexes(mon, 1, idx)
End Sub

'   �Z�̃C���f�b�N�X�̃Z�b�g
Private Sub setAtkIndexes(ByRef mon As Monster, ByVal atkClass As Integer, _
                Optional ByVal idx As Integer = 0)
    If idx <= mon.atkNum Then
        With mon.atkIndex(atkClass)
            .lower = -1
            .upper = -1
            .selected = -1
        End With
    Else
        With mon.atkIndex(atkClass)
            .lower = mon.atkNum
            .upper = idx - 1
            .selected = mon.atkNum
        End With
        mon.atkNum = idx
    End If
End Sub

'   �`���[�W�����̌v�Z
Private Function calcChargeCount(ByRef mon As Monster) As Boolean
    Dim charge As Double
    If mon.atkIndex(0).selected < 0 Or mon.atkIndex(1).selected < 0 Then
        mon.chargeCount = 1.7E+308
        calcChargeCount = False
    End If
    charge = mon.attacks(mon.atkIndex(0).selected).normal.charge
    If charge = 0 Then
        mon.chargeCount = 1.7E+308
        calcChargeCount = False
    Else
        calcChargeCount = True
        With mon.attacks(mon.atkIndex(1).selected).special
            mon.chargeCount = .rsvVol / charge
            '   �Q�[�W1�{�̏ꍇ�͗]���̃`���[�W�͖��ʂɂȂ邽�߁A�`���[�W�J�E���g�͐؂�グ��
            If mon.mode = C_IdGym And .rsvNum = 1 Then
                mon.chargeCount = WorksheetFunction.RoundUp(mon.chargeCount, 0)
            End If
        End With
    End If
End Function

'   �m���J�ڂ̂��߂̕ϐ��̎擾
Public Function initEffectTrans(ByRef effect As AttackEffect, _
            ByRef attr As Object, _
            Optional ByVal considerEffect As Boolean = True) As Boolean
    Dim arr As Variant
    Dim i, idx, col, stageNum As Long
    Dim ifctr(), istage() As Double
    
    '   �R���e�i�̏�����
    arr = Array(Nothing, Nothing, Nothing, Nothing)
    effect.factor = arr
    effect.stages = arr
    effect.expect = Array(1#, 1#, 1#, 1#)
    initEffectTrans = False
    If Not considerEffect Or attr(ATK_Effect) = "" Then Exit Function
    '   ���ʂ�����ꍇ
    effect.desc = attr(ATK_Effect)
    effect.step = attr(ATK_EffectStep)
    effect.probability = attr(ATK_EffectProb)
    '   �W���ƃX�e�[�W�̏����l��ݒ�
    With Range(R_StatusTransition)  '������ʂ̏㏸���E���~���̕\
        If InStr(effect.desc, C_Up) > 0 Then
            col = 2
        ElseIf InStr(effect.desc, C_Down) > 0 Then
            col = 3
        Else
            MsgBox msgstr(msgNoIdentifier, _
                Array(shSpecialAttack.name, attr(ATK_Name), ATK_Effect, C_Up & C_Down))
            Exit Function
        End If
        stageNum = .rows.count
        ReDim ifctr(stageNum - 1), istage(stageNum - 1)
        For i = 1 To stageNum
            ifctr(i - 1) = .cells(i, col)
            istage(i - 1) = 0
        Next
        istage(0) = 1
    End With
    '   ���ʂ̐悪�������G���ŃC���f�b�N�X��ݒ�
    If InStr(effect.desc, C_Self) > 0 Then
        idx = IDX_SelfAtk
    ElseIf InStr(effect.desc, C_Enemy) > 0 Then
        idx = IDX_EnemyAtk
    End If
    '   �u�U���v������
    If InStr(effect.desc, C_Attack) > 0 Then
        effect.factor(idx) = ifctr
        effect.stages(idx) = istage
    End If
    '   �u�h��v������
    If InStr(effect.desc, C_Defense) > 0 Then
        effect.factor(idx + 1) = ifctr
        effect.stages(idx + 1) = istage
    End If
    initEffectTrans = True
End Function

'   HP�Ƃ킴�̓�����ʂ����Z�b�g
Public Sub resetHpAndEffectTrans(ByRef mon As Monster)
    Dim i, idx As Integer
    '   HP�t��
    mon.curHP = mon.fullHP
    mon.phase = 0
    '   �^�����_���[�W�̃��Z�b�g
    mon.given(0).time = 0
    mon.given(0).damage = 0
    mon.given(1).time = 0
    mon.given(1).damage = 0
    '   �Q�[�W�킴
    With mon.attacks(mon.atkIndex(1).selected)
        .special.curVol = 0 '   �Q�[�W���O��
        If Not .isEffect Then Exit Sub
        '   ������ʂ�����ꍇ�͏�����
        With .effect
            For idx = 0 To UBound(.stages)
                If IsArray(.stages(idx)) Then
                    .stages(idx)(0) = 1
                    For i = 1 To UBound(.stages(idx))
                        .stages(idx)(i) = 0
                    Next
                End If
            Next
            .expect = Array(1, 1, 1, 1)
        End With
    End With
End Sub

'   �m���J��
Private Function stochTrans(ByRef effect As AttackEffect)
    Dim i, j, above, stageLimit, trLimit As Long
    Dim sum, prsum As Double
    Dim nvar() As Double
    trLimit = UBound(effect.factor)
    For i = 0 To trLimit
        effect.expect(i) = 1#   '   �f�t�H���g
        If IsArray(effect.factor(i)) Then
            stageLimit = UBound(effect.factor(i))
            ' ���̏�Ԃ��v�Z����
            ReDim nvar(stageLimit)
            For j = 0 To stageLimit
                above = j + effect.step
                If above <= stageLimit Then
                    nvar(j) = nvar(j) + effect.stages(i)(j) * (1 - effect.probability)
                    nvar(above) = nvar(above) + effect.stages(i)(j) * effect.probability
                Else
                    nvar(j) = nvar(j) + effect.stages(i)(j)
                End If
            Next
            effect.stages(i) = nvar
            effect.expect(i) = 0: sum = 0
            For j = 0 To stageLimit
                sum = sum + nvar(j)
                effect.expect(i) = effect.expect(i) + nvar(j) * effect.factor(i)(j)
            Next
            If sum - 1# > 0.0001 Then
                MsgBox ("Something wrong at stochastic transition.")
            End If
        End If
    Next
End Function

'   �̂̎擾
Public Function getIndividual(ByVal identifier As Variant, _
                ByRef mon As Monster, _
                Optional ByVal prediction As Boolean = False)
    Dim ind As Object
    Dim nickname As String
    Dim snum As Long
    Dim indAttr As Variant
    Dim PL As Double
    
    getIndividual = False
    '   �K�v�ȃp�����[�^
    indAttr = Array( _
        IND_Species, IND_indATK, IND_indDEF, IND_indHP, IND_PL, IND_prPL)
    '   �p�����[�^�̎擾
    If IsObject(identifier) Then    ' Range
        nickname = getColumn(IND_Nickname, identifier).Text
        indAttr = getRowValues(identifier, indAttr)
    Else  ' �j�b�N�l�[��
        nickname = identifier
        indAttr = seachAndGetValues( _
                identifier, IND_Nickname, shIndividual, indAttr)
    End If
    '   ���͓r�����`�F�b�N
    If IsEmpty(indAttr(1)) Or IsEmpty(indAttr(2)) _
            Or IsEmpty(indAttr(3)) Or IsEmpty(indAttr(4)) Then
        Exit Function
    End If
    If prediction Then PL = indAttr(5) Else PL = indAttr(4)
    Call clearMonster(mon, nickname, indAttr(0), PL, _
                indAttr(1), indAttr(2), indAttr(3))
    getIndividual = True
End Function

'   �̂̂킴�̎擾
Public Sub setIndividualAttacks(ByRef mon As Monster, _
            ByVal mode As Integer, _
            Optional ByVal prediction As Integer = 0, _
            Optional ByVal cel As Range = Nothing, _
            Optional ByVal atkDelay As Double = 0, _
            Optional ByVal considerEffect As Boolean = False)
    Dim atkNames As Variant
    Dim natks, satks, nflags, sflags As Variant
    '   �킴���̎擾
    atkNames = Array( _
        IND_NormalAtk, IND_SpecialAtk1, IND_SpecialAtk2, _
        IND_TargetNormalAtk, IND_TargetSpecialAtk)
    If Not cel Is Nothing Then    ' �Z���Ŏ擾
        atkNames = getRowValues(cel, atkNames)
    Else  ' �j�b�N�l�[���Ō���
        atkNames = seachAndGetValues( _
                mon.nickname, IND_Nickname, shIndividual, atkNames)
    End If
    mon.mode = mode
    mon.atkNum = 0
    If prediction = 1 Then
        ReDim mon.attacks(1)
        natks = Array(atkNames(3))
        nflags = Array(2)
        satks = Array(atkNames(4))
        sflags = Array(2)
    ElseIf prediction = 2 Then
        ReDim mon.attacks(4)
        natks = Array(atkNames(3), atkNames(0))
        nflags = Array(2, 1)
        satks = Array(atkNames(4), atkNames(1), atkNames(2))
        sflags = Array(2, 1, 1)
    Else
        ReDim mon.attacks(2)
        natks = Array(atkNames(0))
        nflags = Array(1)
        satks = Array(atkNames(1), atkNames(2))
        sflags = Array(1, 1)
    End If
    mon.atkNum = 0
    Call setNormalAttacks(mon, natks, atkDelay, nflags)
    Call setSpecialAttacks(mon, satks, atkDelay, considerEffect, sflags)
End Sub

'   �p�����[�^���w�肵�Č̂��擾
Public Sub getMonster(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal PL As Double = 40, _
                    Optional ByVal indATK As Integer = 15, _
                    Optional ByVal indDEF As Integer = 15, _
                    Optional ByVal indHP As Integer = 15, _
                    Optional ByVal defHP As Integer = 0)
    '   �f�t�H���g�l
    If species = "" Then
        Call clearMonster(mon)
        mon.atkPower = 100
        mon.defPower = 100
        mon.hpPower = 100
        mon.fullHP = 100
        mon.curHP = mon.fullHP
        Exit Sub
    End If
    Call clearMonster(mon, "", species, PL, indATK, indDEF, indHP, defHP)
End Sub

'   �p�����[�^���w�肵�Č̂��擾
Public Sub getMonsterByPower(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal atk As Double = 100, _
                    Optional ByVal def As Double = 100, _
                    Optional ByVal hp As Double = 100)
    Dim attr As Variant
    Dim CPM As Double
    Call clearMonster(mon, "", species)
    mon.atkPower = atk
    mon.defPower = def
    mon.hpPower = hp
    mon.fullHP = Fix(hp)
    mon.curHP = mon.fullHP
    If species <> "" Then
        attr = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
        '   PL�����炵�āA�}�C�i�X���Ȃ��l��������
        For mon.PL = 40 To 1 Step -0.5
            CPM = getCPM(mon.PL)
            mon.indATK = atk / CPM - attr(0)
            mon.indDEF = def / CPM - attr(1)
            mon.indHP = hp / CPM - attr(2)
            If mon.indATK >= 0 And mon.indDEF >= 0 And mon.indHP >= 0 Then Exit For
        Next
        '   �v�Z�s�\
        If mon.PL < 1 Then
            mon.PL = 0
            mon.indATK = 0
            mon.indDEF = 0
            mon.indHP = 0
        End If
        mon.CP = Fix(atk * Sqr(def) * Sqr(hp) / 10)
    End If
End Sub

'   �p�����[�^���w�肵�Č̂��擾
Public Sub getMonsterByCpHp(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal CP As Integer = 1000, _
                    Optional ByVal hp As Integer = 100)
    Dim attr As Variant
    Dim CPd, HPd, CPM, ADmax, AD, CPpHP As Double
    Dim a, b, c, p, q, u, v, ind As Double
    Call clearMonster(mon, "", species)
    If species <> "" Then
        attr = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
    Else
        attr = Array(112, 112, 112)
    End If
    CPd = CP + 0.5
    HPd = hp + 0.5
    ADmax = (attr(0) + 15) * Sqr(attr(1) + 15)
    CPpHP = 10 * CPd / Sqr(HPd)
    '   PL��AD�̌���
    CPM = getCPM(40)
    AD = CPpHP / CPM ^ 1.5 '   AD40
    If AD >= ADmax Then
        mon.PL = 40
    Else
        For mon.PL = 40 To 1.5 Step -0.5
            CPM = getCPM(mon.PL - 0.5)
            AD = CPpHP / CPM ^ 1.5
            If ADmax < AD Then Exit For
        Next
    End If
    CPM = getCPM(mon.PL)
    AD = CPpHP / CPM ^ 1.5
    '   ATK, DEF
    a = 2 * attr(0) + attr(1): b = attr(0) ^ 2 + 2 * attr(0) * attr(1)
    c = attr(0) ^ 2 * attr(1) - AD ^ 2
    p = (b - a ^ 2 / 3) / 3: q = (c + 2 * a ^ 3 / 27 - a * b / 3) / 2
    u = WorksheetFunction.power(-q + Sqr(q ^ 2 + p ^ 2), 1 / 3)
    v = WorksheetFunction.power(-q - Sqr(q ^ 2 + p ^ 2), 1 / 3)
    ind = u + v - a / 3
    mon.indATK = ind: mon.atkPower = (attr(0) + ind) * CPM
    mon.indDEF = ind: mon.defPower = (attr(1) + ind) * CPM
    mon.indHP = HPd / CPM - attr(2): mon.hpPower = HPd
    mon.CP = CP: mon.fullHP = hp: mon.curHP = hp
End Sub

'   �킴�̐ݒ�
Public Sub setAttacks(ByVal mode As Integer, ByRef mon As Monster, _
            Optional ByVal normalAtk As Variant = "", _
            Optional ByVal specialAtk As Variant = "", _
            Optional ByVal atkDelay As Double = 0, _
            Optional ByVal considerEffect As Boolean = False)
    mon.mode = mode
    mon.atkNum = 0
    If Not IsArray(normalAtk) Then normalAtk = Array(normalAtk)
    If Not IsArray(specialAtk) Then normalAtk = Array(specialAtk)
    Call setNormalAttacks(mon, normalAtk, atkDelay)
    Call setSpecialAttacks(mon, specialAtk, atkDelay, considerEffect)
End Sub

'   �_�~�[�̃����X�^�[
Public Function getDummyMonster(ByRef mon As Monster, _
                ByRef param As Object, _
                Optional ByVal mode As Integer = 0, _
                Optional ByVal atkDelay As Double = 0, _
                Optional ByVal types As Variant = Nothing)
    Call clearMonster(mon)
    If IsArray(types) Then
        Call types2idx(types)
        mon.itype(0) = types(0)
        mon.itype(1) = types(1)
    End If
    mon.mode = mode
    mon.atkPower = param(DM_AtkPower)
    mon.defPower = param(DM_DefPower)
    mon.fullHP = param(DM_HP)
    mon.curHP = mon.fullHP
    ReDim mon.attacks(1)
    With mon.attacks(0)
        .name = ""
        .class = C_IdNormalAtk
        .itype = 0
        If mode = C_IdGym Then
            .power = param(DM_GymNAtkPower)
            .idleTime = param(DM_GymNAtkIdleTime) + atkDelay
            .normal.charge = param(DM_GymNAtkCharge)
        ElseIf mode = C_IdMtc Then
            .power = param(DM_MtcNAtkPower)
            .idleTime = param(DM_MtcNAtkIdleTurn) + atkDelay
            .normal.charge = param(DM_MtcNAtkCharge)
        End If
    End With
    With mon.attacks(1)
        .name = ""
        .class = C_IdSpecialAtk
        .itype = 0
        If mode = C_IdGym Then
            .power = param(DM_GymSAtkPower)
            .idleTime = param(DM_GymSAtkIdleTime) + atkDelay
            .special.rsvNum = param(DM_GymSAtkGuageNum)
            .special.rsvVol = 100# / .special.rsvNum
        ElseIf mode = C_IdMtc Then
            .power = param(DM_MtcSAtkPower)
            .idleTime = 1 + atkDelay
            .special.rsvNum = 1
            .special.rsvVol = param(DM_MtcSAtkGuageVol)
        End If
        .isEffect = False
    End With
    With mon.atkIndex(0)
        .lower = 0
        .upper = 0
        .selected = 0
    End With
    With mon.atkIndex(1)
        .lower = 1
        .upper = 1
        .selected = 1
    End With
End Function
    

'   �N���A
Public Sub clearMonster(ByRef mon As Monster, _
        Optional ByVal nickname As String = "", _
        Optional ByVal species As String = "", _
        Optional ByVal PL As Double = 40, _
        Optional ByVal indATK As Integer = 15, _
        Optional ByVal indDEF As Integer = 15, _
        Optional ByVal indHP As Integer = 15, _
        Optional ByVal defHP As Integer = 0)
    Dim types As Variant
    mon.nickname = nickname
    mon.species = species
    If nickname <> "" Then mon.logName = nickname Else mon.logName = species
    mon.itype(0) = 0
    mon.itype(1) = 0
    If species <> "" Then
        types = getSpcAttrs(species, Array(SPEC_Type1, SPEC_Type2))
        mon.itype(0) = getTypeIndex(types(0))
        mon.itype(1) = getTypeIndex(types(1))
    End If
    mon.PL = PL
    mon.indATK = indATK
    mon.indDEF = indDEF
    mon.indHP = indHP
    '   ���̑��N���A
    mon.fullHP = 0
    mon.curHP = 0
    mon.atkPower = 0
    mon.defPower = 0
    mon.hpPower = 0
    mon.CP = 0
    mon.chargeCount = 0
    mon.clock = 0
    mon.phase = 0
    mon.mode = -1
    mon.atkNum = 0
    mon.atkIndex(0).selected = -1
    mon.atkIndex(1).selected = -1
    mon.given(0).time = 0
    mon.given(0).damage = 0
    mon.given(1).time = 0
    mon.given(1).damage = 0
    Call calcMonPowers(mon, defHP)
End Sub

'   PL�A�̒l���e�p���[�̌v�Z
Public Sub calcMonPowers(mon As Monster, _
                    Optional ByVal defHP As Integer = 0)
    Dim attrs As Variant
    Dim CPM As Double
    If mon.species = "" Then Exit Sub
    attrs = getSpcAttrs(mon.species, Array("ATK", "DEF", "HP"))
    CPM = getCPM(mon.PL)
    mon.atkPower = (mon.indATK + attrs(0)) * CPM
    mon.defPower = (mon.indDEF + attrs(1)) * CPM
    mon.fullHP = defHP
    If defHP = 0 Then
        mon.hpPower = (mon.indHP + attrs(2)) * CPM
        mon.fullHP = Fix(mon.hpPower)
    Else
        mon.fullHP = defHP
        mon.hpPower = defHP + 0.5
    End If
    mon.curHP = mon.fullHP
    mon.CP = Fix(mon.atkPower * Sqr(mon.defPower) * Sqr(mon.hpPower) / 10)
End Sub

'   �_���[�W���܂Ƃ߂Čv�Z
Private Sub calcDamages( _
            ByRef offence As Monster, ByRef deffence As Monster, _
            Optional ByVal isAna As Boolean = True, _
            Optional ByVal weather As Integer = 0)
    Dim effect As Double
    Dim atk As Variant
    '   ������ʌv�Z�B�U����
    With offence.attacks(offence.atkIndex(1).selected)
        effect = 1#
        If .isEffect Then
            effect = .effect.expect(IDX_SelfAtk) / .effect.expect(IDX_EnemyDef)
        End If
    End With
    '   ������ʌv�Z�B�h�䑤
    If deffence.atkIndex(1).selected >= 0 Then
        With deffence.attacks(deffence.atkIndex(1).selected)
            If .isEffect Then
                effect = effect * .effect.expect(IDX_EnemyAtk) / .effect.expect(IDX_SelfDef)
            End If
        End With
    End If
    '   ��͂�false�ł��A������ʂ������true�ɂ��ď����_�ȉ�������
    If isAna = False And effect <> 1 Then
        isAna = True
    End If
    Call calcADamage(offence.atkIndex(0).selected, offence, deffence, isAna, weather, effect)
    Call calcADamage(offence.atkIndex(1).selected, offence, deffence, isAna, weather, effect)
End Sub

'   �_���[�W���v�Z
Public Function calcADamage(ByRef atkIdx As Integer, _
            ByRef offence As Monster, ByRef deffence As Monster, _
            Optional ByVal isAna As Boolean = True, _
            Optional ByVal weather As Integer = 0, _
            Optional ByVal effect As Double = 1#) As Double
    Dim fctr As Double
    If atkIdx < 0 Then Exit Function
    With offence.attacks(atkIdx)
        If deffence.defPower = 0 Then
            Debug.Print "defPower is 0. " & deffence.nickname
            .damage = 0
            Exit Function
        End If
        fctr = getFactor(offence.mode, offence.itype, .itype, deffence.itype, weather)
        fctr = fctr * effect
        .damage = (offence.atkPower / deffence.defPower * .power * fctr * 0.5) + 1
        If Not isAna Then .damage = Fix(.damage)
        calcADamage = .damage
    End With
End Function

'   �W���̎擾
Public Function getFactor(ByVal mode As Integer, _
                ByVal selfTypes As Variant, _
                ByVal atkType As Integer, _
                ByVal enemyTypes As Variant, _
                ByVal weather As Integer) As Double
    If Not IsArray(InterTypeFactorTable) Then Call makeInfluenceCache
    getFactor = 1#
    '   �^�C�v��v
    If atkType > 0 Then
        If atkType = selfTypes(0) Or atkType = selfTypes(1) Then
            getFactor = TypeMatchFactorValue
        End If
    End If
    '   �^�C�v����
    getFactor = getFactor * InterTypeFactorTable(atkType, enemyTypes(0)) _
                * InterTypeFactorTable(atkType, enemyTypes(1))
    '   ���̑�
    If mode = C_IdGym Then   ' �W���o�g���ł͓V��u�[�X�g
        getFactor = getFactor * WeatherFactorTable(atkType, weather)
    ElseIf mode = C_IdMtc Then   ' �g���[�i�[�o�g����1.3�{
        getFactor = getFactor * MatchBattleFactorValue
    End If
End Function

'   ���փV�[�g�Ȃǂ̕\�̃L���b�V���̍쐬
Public Sub makeInfluenceCache()
    InterTypeFactorTable = makeInterTypeTable
    WeatherFactorTable = makeWeatherFactorTable
    TypeMatchFactorValue = Range(R_TypeMatchFactor).Value
    MatchBattleFactorValue = Range(R_MtcBtlFactor).Value
    ChargeByDamage = Range(R_ChargeByDamage).Value
End Sub

'   ���֕\�̍쐬
Public Function makeInterTypeTable() As Variant
    Dim tbl() As Double
    Dim i, j, n As Integer
    Dim mark As String
    n = typesNum()
    ReDim tbl(n, n)
    tbl(0, 0) = 1
    For i = 1 To n
        tbl(0, i) = 1: tbl(i, 0) = 1
        For j = 1 To n
            mark = Range(R_InterTypeInflu).cells(i, j).Text
            If mark <> "" Then
                tbl(i, j) = WorksheetFunction.VLookup( _
                        mark, Range(R_interTypeFactor), 3, False)
            Else
                tbl(i, j) = 1
            End If
        Next
    Next
    makeInterTypeTable = tbl
End Function

'   �V��u�[�X�g�W���̃L���b�V�������
Private Function makeWeatherFactorTable()
    Dim tbl() As Double
    Dim ti, wi, tn, wn As Integer
    Dim boost As String
    tn = typesNum()
    wn = Range(R_WeatherTable).rows.count
    ReDim tbl(tn, wn)
    For ti = 0 To tn
        For wi = 0 To wn
            tbl(ti, wi) = 1
        Next
        If ti > 0 Then
            boost = Range(R_WeatherBoost).cells(ti, 1).Text
            wi = getWeatherIndex(boost)
            tbl(ti, wi) = Range(R_WeatherFactor).Value
        End If
    Next
    makeWeatherFactorTable = tbl
End Function

'   �^�C�v��v�W��
Private Function atkTypeFactor(ByVal self As Variant, ByVal atkType As Variant) As Double
    atkTypeFactor = 1
    Call type2idx(atkType)
    If atkType <> 0 Then
        Call types2idx(self)
        If self(0) = atkType Or self(1) = atkType Then
            atkTypeFactor = Range(R_TypeMatchFactor).Value
        End If
    End If
End Function

'   ����Ƃ̑����̌W��
Public Function interTypeFactor(ByVal atkType As Integer, _
                ByVal enemy As Variant) As Double
    If Not IsArray(InterTypeFactorTable) Then Call makeInfluenceCache
    interTypeFactor = InterTypeFactorTable(atkType, enemy(0)) * InterTypeFactorTable(atkType, enemy(1))
End Function

'   �V��u�[�X�g�̌W��
Private Function weatherFactor(ByVal atkType As Integer, ByVal weather As Integer) As Double
    If Not IsArray(WeatherFactorTable) Then Call makeInfluenceCache
    WeatherFactorTable = WeatherFactorTable(atkType, weather)
End Function

Public Sub type2idx(ByRef tp As Variant)
    If Not IsNumeric(tp) Then tp = getTypeIndex(tp)
End Sub

Public Sub types2idx(ByRef tps As Variant)
    If Not IsNumeric(tps(0)) Then tps(0) = getTypeIndex(tps(0))
    If Not IsNumeric(tps(1)) Then tps(1) = getTypeIndex(tps(1))
End Sub


'   KT�܂���KTR�ɂ�郉���N�𓾂�
'   �Ԃ�l�̔z��̓Y����0����B��v�f�͔z��ŁA���̊e�v�f�͈ȉ��B
'   0:KTR, 1:KT, 2:�j�b�N�l�[��, 3:�ʏ�Z,
'   4:�ʏ�ZtDPS, 5:�Q�[�W�Z, 6:�Q�[�W�ZtDPS, 7:cDPS
Public Function getKtRank(ByVal rankNum As Long, _
        ByRef enemy As Monster, _
        Optional ByVal isKTR As Boolean = True, _
        Optional ByVal upperCP As Long = 0, _
        Optional ByVal lowerCP As Long = 0, _
        Optional ByVal prediction As Boolean = False, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal selfAtkDelay As Double = 0) As Variant
    Dim ktrs, minKtrs, rank(), vtmp As Variant
    Dim cel As Range
    Dim ri, nai, sai As Long
    Dim predict As Integer
    Dim CP As Long
    Dim self As Monster
    Dim spesifiedEnemyAtk As Boolean
    
    If Not IsNumeric(weather) Then weather = getWeatherIndex(weather)
    If prediction Then predict = 2 Else predict = 0
    With enemy  '   �G�̋Z�����肳��Ă��邩
        spesifiedEnemyAtk = (.atkIndex(0).lower = .atkIndex(0).upper _
                And .atkIndex(1).lower = .atkIndex(1).upper)
    End With
    ReDim rank(rankNum)
    For Each cel In shIndividual.ListObjects(1).ListColumns(IND_CP).DataBodyRange
        dspProgress
        '   CP�̃`�F�b�N�B�����E������L����CP�����͈̔͂𒴂��Ă����玟��
        CP = cel.Value
        If (lowerCP > 0 And CP < lowerCP) Or (upperCP > 0 And upperCP < CP) Then GoTo Continue
        '   �̂̐ݒ�
        Call getIndividual(cel, self, prediction)
        Call setIndividualAttacks(self, enemy.mode, predict, cel, selfAtkDelay, _
                                    (enemy.mode = C_IdMtc))
        If self.PL = 0 Or self.atkIndex(0).lower < 0 Or self.atkIndex(1).lower < 0 Then GoTo Continue
        minKtrs = Array(1E+107, 1E+107)
        If self.atkIndex(0).lower < 0 Or self.atkIndex(1).lower < 0 Then GoTo Continue
        '   �̂̒ʏ�E�Q�[�W�Z�ŁA�ŏ���KTR�܂���KT�̂��̂𓾂�
        For nai = self.atkIndex(0).lower To self.atkIndex(0).upper
            self.atkIndex(0).selected = nai
            For sai = self.atkIndex(1).lower To self.atkIndex(1).upper
                self.atkIndex(1).selected = sai
                '   �G�̋Z�����肳��Ă�����KTR���擾
                If spesifiedEnemyAtk Then
                    Call resetHpAndEffectTrans(self)
                    Call resetHpAndEffectTrans(enemy)
                    ktrs = simBattle(self, enemy, weather, True)
                Else    '   ���肳��Ă��Ȃ��̂ŕ��ς𓾂�
                    ktrs = getAveKTR(self, enemy, weather)
                End If
                '   �ŏ���minKtrs�ɋL�^
                If (isKTR And minKtrs(0) > ktrs(0)) _
                        Or (Not isKTR And minKtrs(1) > ktrs(1)) Then
                    ReDim Preserve ktrs(7)
                    ktrs(3) = nai
                    ktrs(5) = sai
                    '   KTR���擾�����ꍇ�͐��`
                    If spesifiedEnemyAtk Then
                        With self
                            ktrs(4) = .given(0).damage / .given(0).time
                            If .given(1).time > 0 Then
                                ktrs(6) = .given(1).damage / .given(1).time
                            Else
                                ktrs(6) = ""
                            End If
                            ktrs(7) = (.given(0).damage + .given(1).damage) _
                                    / (.given(0).time + .given(1).time)
                        End With
                    End If
                    minKtrs = ktrs
                End If
            Next
        Next
        '   ���̂�KTR(KT)���ŏ��̂��̂̐��`
        minKtrs(2) = self.nickname
        If IsNumeric(minKtrs(3)) Then
            minKtrs(3) = self.attacks(minKtrs(3)).name
            minKtrs(5) = self.attacks(minKtrs(5)).name
        End If
        rank(rankNum) = minKtrs
        '   �����N�̍X�V
        ri = rankNum - 1
        Do While ri >= 0
            If Not IsEmpty(rank(ri)) Then
                If (isKTR And rank(ri)(0) <= rank(ri + 1)(0)) _
                Or (Not isKTR And rank(ri)(1) <= rank(ri + 1)(1)) Then
                    Exit Do
                End If
            End If
            vtmp = rank(ri): rank(ri) = rank(ri + 1): rank(ri + 1) = vtmp
            ri = ri - 1
        Loop
Continue:
    Next
    ReDim Preserve rank(rankNum - 1)
    getKtRank = rank
End Function


'   �G�̋Z�𑍓��肵�āAKT�AKTR�̕��ϒl���擾�B
'   �߂�l�͔z��B���̗v�f�͔z���
'   0:KRT, 1:��KT, 4:�ʏ�킴tDPS, 6:�Q�[�W�킴tDPA, 7:cDPS
Public Function getAveKTR( _
        ByRef self As Monster, ByRef enemy As Monster, _
        Optional ByVal weather As String = "") As Variant
    Dim sumKtr(7) As Variant
    Dim ktrs As Variant
    Dim nai, sai, cnt, i As Long
    '   �N���A
    For i = 0 To UBound(sumKtr)
        If i = 2 Or i = 3 Or i = 5 Then
            sumKtr(i) = ""
        Else
            sumKtr(i) = 0#
        End If
    Next
    With enemy
        For nai = .atkIndex(0).lower To .atkIndex(0).upper
            .atkIndex(0).selected = nai
            For sai = .atkIndex(1).lower To .atkIndex(1).upper
                .atkIndex(1).selected = sai
                Call resetHpAndEffectTrans(self)
                Call resetHpAndEffectTrans(enemy)
                ktrs = simBattle(self, enemy, weather, True)
                sumKtr(0) = sumKtr(0) + ktrs(0)
                sumKtr(1) = sumKtr(1) + ktrs(1)
                With self
                    sumKtr(4) = sumKtr(4) _
                        + (.given(0).damage / .given(0).time)
                    If .given(1).time > 0 Then
                        sumKtr(6) = sumKtr(6) _
                            + (.given(1).damage / .given(1).time)
                    End If
                    sumKtr(7) = sumKtr(7) _
                        + (.given(0).damage + .given(1).damage) _
                            / (.given(0).time + .given(1).time)
                End With
                cnt = cnt + 1
            Next
        Next
    End With
    For i = 0 To UBound(sumKtr)
        If IsNumeric(sumKtr(i)) Then
            sumKtr(i) = sumKtr(i) / cnt
        End If
    Next
    getAveKTR = sumKtr
End Function


