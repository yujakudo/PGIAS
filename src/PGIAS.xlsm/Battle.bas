Attribute VB_Name = "Battle"

Option Explicit

'   ゲージの最大量
Const RSV_MAX As Double = 100#

'   特殊効果のインデックス
Const IDX_SelfAtk As Integer = 0
Const IDX_SelfDef As Integer = 1
Const IDX_EnemyAtk As Integer = 2
Const IDX_EnemyDef As Integer = 3

Dim InterTypeFactorTable As Variant
Dim WeatherFactorTable As Variant
Dim TypeMatchFactorValue As Double
Dim MatchBattleFactorValue As Double
Dim ChargeByDamage As Double

'   通常わざ
Public Type NormalAttack
    idleTime As Double
    charge As Integer
End Type

'   効果
Public Type AttackEffect
    desc As String
    probability As Double
    step As Integer
    factor As Variant
    stages As Variant
    expect As Variant
End Type

'   ゲージわざ
Public Type SpecialAttack
    rsvNum As Integer
    rsvVol As Double
    curVol As Double
End Type

'   わざ
Public Type Attack
    name As String
    itype As Integer
    Power As Double
    damage As Double
    idleTime As Double
    class As Integer
    normal As NormalAttack
    special As SpecialAttack
    isEffect As Boolean
    effect As AttackEffect
    flag As Integer
End Type

'   わざのインデックス
Public Type AttackIndex
    selected As Integer
    lower As Integer
    upper As Integer
End Type

'   相手に与えたダメージの累積
Public Type GivenDamage
    time As Double
    damage As Double
End Type

'   ポケモン
Public Type Monster
    nickname As String
    species As String
    logName As String
    itype(1) As Integer
    PL As Double
    indATK As Long
    indDEF As Long
    indHP As Long
    atkPower As Double
    defPower As Double
    hpPower As Double
    fullHP As Long
    curHP As Double
    CP As Long
    attacks() As Attack
    atkIndex(1) As AttackIndex
    given(1) As GivenDamage
    atkNum As Integer
    chargeCount As Double
    clock As Double
    mode As Integer
    phase As Integer
End Type

'   技の評価値
Public Type AttackParam
    name As String
    class As Integer
    damage As Double
    damageEfc As Double
    chargeEfc As Double
End Type

'   CDPSの情報
Public Type CDpsSet
    natk As String
    satk As String
    cDPS As Double
    Cycle As Double
End Type

'   ランキング情報内の技情報
Public Type AtkParam
    idx As Integer
    name As String
    tDPS As Double
    stDPS As Double
End Type

'   ランキング情報
Public Type SimInfo
    nickname As String
    PL As Double
    KTR As Double
    KT As Double
    cDPS As Double
    scDPS As Double
    Cycle As Double
    Attack(1) As AtkParam
    flag As Integer
End Type

'   ランキング情報の配列
Public Type KtRank
    rank() As SimInfo
End Type


'   ダメージの計算
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

'   ダメージの計算（解析用）
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

'   選択されている技のインデックスの取得
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

'   cDPSの計算
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
    Dim cdpss As CDpsSet
    Call getMonster(self, species, PL, indATK)
    Call setAttacks(mode, self, natkName, satkName)
    Call getMonster(enemy, enemySpecies, enemyPL, 0, enemyIndDef)
    cdpss = calcCDPS(self, enemy, isAna, weather)
    getCDPS = cdpss.cDPS
End Function

'   cDPSの算出
Public Function calcCDPS(ByRef self As Monster, ByRef enemy As Monster, _
                Optional ByVal isAna As Boolean = False, _
                Optional ByVal weather As Variant = 0, _
                Optional ByVal ignoreFactor = False) As CDpsSet
    Dim damage, period As Double
    Dim atkIdx As Variant
    
    atkIdx = getAttackIndex(self)
    With self
        calcCDPS.natk = .attacks(atkIdx(0)).name
        calcCDPS.satk = .attacks(atkIdx(1)).name
    End With
    If Not IsNumeric(weather) Then weather = getWeatherIndex(weather)
    Call calcDamages(self, enemy, isAna, weather, ignoreFactor)
    If calcChargeCount(self) Then
        With self
            damage = .attacks(atkIdx(0)).damage * self.chargeCount + .attacks(atkIdx(1)).damage
            period = .attacks(atkIdx(0)).idleTime * self.chargeCount + .attacks(atkIdx(1)).idleTime
        End With
        If period = 0 Then
            Debug.Print "Period is 0 in calcCDPS. " & self.nickname
        End If
        calcCDPS.cDPS = damage / period
        calcCDPS.Cycle = period
    End If
End Function

'   攻撃解析パラメータの取得
'   戻り値の配列の要素は1から順に、ダメージ、ダメージ効率、チャージ効率（通常わざのみ）
Public Function getAtkAna( _
        ByRef AtkParam() As AttackParam, _
        ByRef cdpss() As CDpsSet, _
        ByVal mode As Integer, _
        ByVal species As String, _
        Optional ByVal withLimited As Boolean = True, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal PL As Double = 40, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal enemySpecies As String = "", _
        Optional ByVal enemyPL As Double = 40, _
        Optional ByVal enemyIndDef As Long = 15)
    Dim atkNames, cDPS As Variant
    Dim self As Monster
    Dim enemy As Monster
    Dim i, j, atkClass, ni, si As Integer
    Dim tmp As CDpsSet
    
    Call getMonster(self, species, PL, indATK)
    If enemySpecies = "" Then
        Call getMonsterByPower(enemy)
    Else
        Call getMonster(enemy, enemySpecies, enemyPL, 0, enemyIndDef)
    End If
    '   すべての技を設定
    atkNames = getAtkNames(species, False, withLimited)
    Call setAttacks(mode, self, atkNames(0), atkNames(1))
    ReDim AtkParam(self.atkNum - 1)
    '   クラスごと、すべての技のパラメータを計算
    For i = 0 To self.atkNum - 1
        AtkParam(i) = getAtkParam(self, enemy, i, weather)
    Next
    '   すべての技組み合わせについてcDSPを計算。ソートもしておく
    With self
        ReDim cdpss((.atkIndex(0).upper - .atkIndex(0).lower + 1) _
                    * (.atkIndex(1).upper - .atkIndex(1).lower + 1) - 1)
        i = 0
        For ni = .atkIndex(0).lower To .atkIndex(0).upper
            self.atkIndex(0).selected = ni
            For si = .atkIndex(1).lower To .atkIndex(1).upper
                self.atkIndex(1).selected = si
                cdpss(i) = calcCDPS(self, enemy, True, weather)
                For j = i - 1 To 0 Step -1
                    If cdpss(j).cDPS < cdpss(j + 1).cDPS Then
                        tmp = cdpss(j): cdpss(j) = cdpss(j + 1): cdpss(j + 1) = tmp
                    End If
                Next
                i = i + 1
            Next
        Next
    End With
End Function


'   攻撃解析パラメータの取得（ダメージ、ダメージ効率、チャージ効率）
Private Function getAtkParam(ByRef self As Monster, ByRef enemy As Monster, _
        ByVal atkIdx As Integer, _
        Optional ByVal weather As Integer = 0) As AttackParam
    Dim atkClass As Integer
    Dim ret() As Double
    Dim attrs() As String
    Dim vals As Variant
    
    With self.attacks(atkIdx)
        atkClass = .class
        '   必要な属性の取得
        If atkClass = C_IdNormalAtk Then   '   通常わざ
            ReDim attrs(1), ret(2)
            '   ダメージ効率の分母とチャージ効率
            If self.mode = C_IdGym Then    '   発動時間, チャージ速度EPS
                attrs(0) = ATK_IdleTime: attrs(1) = ATK_EPS
            ElseIf self.mode = C_IdMtc Then    '   発動ターン, チャージ速度EPT
                attrs(0) = ATK_IdleTurnNum: attrs(1) = ATK_EPT
            End If
        Else    '   ゲージわざ
            ReDim attrs(0), ret(1)
            '   ダメージ効率の分母
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
    '   ダメージ効率
    getAtkParam.damageEfc = getAtkParam.damage / vals(0)
    '   チャージ効率
    If atkClass = C_IdNormalAtk Then getAtkParam.chargeEfc = vals(1)
End Function

'   戦闘シミュレーション
'   返り値の配列の要素は、0:KTR, 1:自KT, 2:敵KT, 3:ログ
Private Function simBattle(ByRef sbjMon As Monster, ByRef objMon As Monster, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal needKTR As Boolean = False, _
        Optional ByVal logCel As Range = Nothing) As Variant
    Dim sbjKT, objKT As Double
    Dim alive(1), isSbjFirst As Boolean
    Dim log As Variant
    Dim clock As Double
    
'    Set logCel = shBattleSim.cells(8, 1)
    sbjMon.clock = 0
    objMon.clock = 0
    sbjKT = 0: objKT = 0
    alive(0) = True: alive(1) = True
    Call calcDamages(sbjMon, objMon, False, weather)
    Call calcDamages(objMon, sbjMon, False, weather)
    isSbjFirst = sbjMon.atkPower >= objMon.atkPower
    '   攻撃を決定しておく
    Call decideAttack(sbjMon)
    Call decideAttack(objMon)
    '   開始時のログ
    If Not logCel Is Nothing Then
        Call logStatus(logCel, 0, sbjMon, objMon)
    End If
    '   どちらかが生きている間
    Do While alive(0) Or alive(1)
        If sbjMon.clock < objMon.clock Or (isSbjFirst And sbjMon.clock = objMon.clock) Then
            clock = sbjMon.clock
            log = hitAttack(sbjMon, objMon, weather)
            '   相手のHPが0以下になったらKTを記録
            If alive(1) And objMon.curHP <= 0 Then
                sbjKT = sbjMon.clock
                alive(1) = False
                If Not needKTR Then Exit Do
            End If
            Call decideAttack(sbjMon)
        Else
            clock = objMon.clock
            log = hitAttack(objMon, sbjMon, weather)
            If alive(0) And sbjMon.curHP <= 0 Then
                objKT = objMon.clock
                alive(0) = False
                If Not needKTR Then Exit Do
            End If
            Call decideAttack(objMon)
        End If
        If Not logCel Is Nothing Then
            Call logStatus(logCel, clock, sbjMon, objMon, log)
        End If
    Loop
    If needKTR Then
        simBattle = Array(sbjKT / objKT, sbjKT, objKT, log)
    Else
        simBattle = Array(0, sbjKT, objKT, log)
    End If
End Function

'   ステータスログ
Private Sub logStatus(ByRef logCel As Range, _
                    ByVal clock As Double, _
                    ByRef sbjMon As Monster, _
                    ByRef objMon As Monster, _
            Optional ByVal damages As Variant = False)
    Dim i As Integer
    logCel.value = clock
    If IsArray(damages) Then
        For i = 0 To 2
            logCel.Offset(0, i + 1).value = damages(i)
        Next
    End If
    With sbjMon
        logCel.Offset(0, 4) = .curHP
        logCel.Offset(0, 5) = .attacks(.atkIndex(1).selected).special.curVol
    End With
    With objMon
        logCel.Offset(0, 6) = .curHP
        logCel.Offset(0, 7) = .attacks(.atkIndex(1).selected).special.curVol
    End With
    Set logCel = logCel.Offset(1, 0)
End Sub

'   次の攻撃の決定
Private Sub decideAttack(ByRef offence As Monster)
    Dim idleTime As Double
    Dim atkIdx As Integer
    
    atkIdx = offence.atkIndex(1).selected
    With offence.attacks(atkIdx)    '   ゲージ技
        '   ゲージが満ちていなければ、通常技を開始
        If .special.curVol < .special.rsvVol Then atkIdx = offence.atkIndex(0).selected
    End With
    '   攻撃準備フェイズ
    With offence.attacks(atkIdx)
        offence.phase = .class + 1
        '   クロック
        offence.clock = offence.clock + .idleTime
    End With
End Sub

'   攻撃
Private Function hitAttack(ByRef offence As Monster, ByRef deffence As Monster, _
        Optional ByVal weather As Integer = 0) As Variant
    Dim charge, damage, idleTime As Double
    Dim name As String
    Dim atkIdx As Integer
    
    atkIdx = offence.atkIndex(offence.phase - 1).selected
    With offence.attacks(atkIdx)
        '   ダメージ
        damage = .damage
        idleTime = .idleTime
        deffence.curHP = deffence.curHP - damage
        With offence.given(.class)
            .damage = .damage + damage
            .time = .time + idleTime
        End With
        '   ダメージによる防御側のチャージ
        With deffence.attacks(deffence.atkIndex(1).selected)
            .special.curVol = .special.curVol + damage * ChargeByDamage
            '   ジムの場合は満タン
            If .special.curVol > RSV_MAX And deffence.mode = C_IdGym Then
               .special.curVol = RSV_MAX
            End If
        End With
        If .class = 0 Then  '   通常技
            '   チャージ
            charge = .normal.charge
            With offence.attacks(offence.atkIndex(1).selected)
                .special.curVol = .special.curVol + charge
                '   ジムでゲージ数1の場合は満タン
                If .special.curVol > .special.rsvVol _
                    And (offence.mode = C_IdGym And .special.rsvNum = 1) Then
                   .special.curVol = .special.rsvVol
                End If
            End With
        Else  '   ゲージ技
            .special.curVol = .special.curVol - .special.rsvVol
            '   特殊効果があればダメージを再計算
            If .isEffect Then
                Call stochTrans(.effect)
                Call calcDamages(offence, deffence, False, weather)
                Call calcDamages(deffence, offence, False, weather)
            End If
        End If
        hitAttack = Array(offence.logName, .name, damage)
    End With
    offence.phase = 0
End Function


'   通常わざの設定。（damageはセットしない）
'   atkNames: 名前、或は設定オブジェクトの配列
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
        With mon.attacks(idx)
            If IsObject(atkNames(i)) Then
                .name = atkNames(i)("name")
                Set nattr = atkNames(i)
            ElseIf atkNames(i) <> "" Then
                .name = atkNames(i)
                Set nattr = getAtkAttrs(C_NormalAttack, .name)
            Else
                GoTo Continue
            End If
            .class = C_IdNormalAtk
            If IsArray(flags) Then .flag = flags(i)
            .itype = getTypeIndex(nattr(ATK_Type))
            If mon.mode = C_IdGym Then
                .Power = nattr(ATK_GymPower)
                .idleTime = nattr(ATK_IdleTime) + atkDelay
                .normal.charge = nattr(ATK_GymCharge)
            ElseIf mon.mode = C_IdMtc Then
                .Power = nattr(ATK_MtcPower)
                .idleTime = nattr(ATK_IdleTurnNum) + atkDelay
                .normal.charge = nattr(ATK_MtcCharge)
            End If
            idx = idx + 1
Continue:
        End With
    Next
    Call setAtkIndexes(mon, 0, idx)
End Sub

'   ゲージわざの設定。（damageはセットしない）
'   atkNames: 名前、或は設定オブジェクトの配列
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
        With mon.attacks(idx)
            If IsObject(atkNames(i)) Then
                .name = atkNames(i)("name")
                Set sattr = atkNames(i)
            ElseIf atkNames(i) <> "" Then
                .name = atkNames(i)
                Set sattr = getAtkAttrs(C_SpecialAttack, .name)
            Else
                GoTo Continue
            End If
            .class = C_IdSpecialAtk
            If IsArray(flags) Then .flag = flags(i)
            .itype = getTypeIndex(sattr(ATK_Type))
            If mon.mode = C_IdGym Then
                .Power = sattr(ATK_GymPower)
                .idleTime = sattr(ATK_IdleTime) + atkDelay
                .special.rsvNum = sattr(ATK_GaugeNumber)
                If .special.rsvNum = 0 Then
                    Debug.Print "Number of Reserver is 0. " & atkNames(i)
                End If
                .special.rsvVol = RSV_MAX / .special.rsvNum
            ElseIf mon.mode = C_IdMtc Then
                .Power = sattr(ATK_MtcPower)
                .idleTime = 1 + atkDelay
                .special.rsvNum = 1
                .special.rsvVol = sattr(ATK_GaugeVolume)
            End If
            .isEffect = initEffectTrans(.effect, sattr, considerEffect)
            idx = idx + 1
Continue:
        End With
    Next
    Call setAtkIndexes(mon, 1, idx)
End Sub

'   技のインデックスのセット
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

'   チャージ発数の計算
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
            '   ゲージ1本の場合は余分のチャージは無駄になるため、チャージカウントは切り上げる
            If mon.mode = C_IdGym And .rsvNum = 1 Then
                mon.chargeCount = WorksheetFunction.RoundUp(mon.chargeCount, 0)
            End If
        End With
    End If
End Function

'   確率遷移のための変数の取得
Public Function initEffectTrans(ByRef effect As AttackEffect, _
            ByRef attr As Object, _
            Optional ByVal considerEffect As Boolean = True) As Boolean
    Dim arr As Variant
    Dim i, idx, col, stageNum As Long
    Dim ifctr(), istage() As Double
    
    '   コンテナの初期化
    arr = Array(Nothing, Nothing, Nothing, Nothing)
    effect.factor = arr
    effect.stages = arr
    effect.expect = Array(1#, 1#, 1#, 1#)
    initEffectTrans = False
    If Not considerEffect Or attr(ATK_Effect) = "" Then Exit Function
    '   効果がある場合
    effect.desc = attr(ATK_Effect)
    effect.step = attr(ATK_EffectStep)
    effect.probability = attr(ATK_EffectProb)
    '   係数とステージの初期値を設定
    With Range(R_StatusTransition)  '特殊効果の上昇率・下降率の表
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
    '   効果の先が自分か敵かでインデックスを設定
    If InStr(effect.desc, C_Self) > 0 Then
        idx = IDX_SelfAtk
    ElseIf InStr(effect.desc, C_Enemy) > 0 Then
        idx = IDX_EnemyAtk
    End If
    '   「攻撃」がある
    If InStr(effect.desc, C_Attack) > 0 Then
        effect.factor(idx) = ifctr
        effect.stages(idx) = istage
    End If
    '   「防御」がある
    If InStr(effect.desc, C_Defense) > 0 Then
        effect.factor(idx + 1) = ifctr
        effect.stages(idx + 1) = istage
    End If
    initEffectTrans = True
End Function

'   HPとわざの特殊効果をリセット
Public Sub resetHpAndEffectTrans(ByRef mon As Monster)
    Dim i, idx As Integer
    '   HPフル
    mon.curHP = mon.fullHP
    mon.phase = 0
    '   与えたダメージのリセット
    mon.given(0).time = 0
    mon.given(0).damage = 0
    mon.given(1).time = 0
    mon.given(1).damage = 0
    '   ゲージわざ
    With mon.attacks(mon.atkIndex(1).selected)
        .special.curVol = 0 '   ゲージを０に
        If Not .isEffect Then Exit Sub
        '   特殊効果がある場合は初期化
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

'   確率遷移
Private Function stochTrans(ByRef effect As AttackEffect)
    Dim i, j, above, stageLimit, trLimit As Long
    Dim sum, prsum As Double
    Dim nvar() As Double
    trLimit = UBound(effect.factor)
    For i = 0 To trLimit
        effect.expect(i) = 1#   '   デフォルト
        If IsArray(effect.factor(i)) Then
            stageLimit = UBound(effect.factor(i))
            ' 次の状態を計算する
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

'   個体の取得
Public Function getIndividual(ByVal identifier As Variant, _
                ByRef mon As Monster, _
                Optional ByVal prediction As Boolean = False) As Boolean
    Dim ind As Object
    Dim nickname As String
    Dim snum As Long
    Dim indAttr As Variant
    Dim PL As Double
    
    getIndividual = False
    '   必要なパラメータ
    indAttr = Array( _
        IND_Species, IND_indATK, IND_indDEF, IND_indHP, IND_PL, IND_prPL)
    '   パラメータの取得
    If IsObject(identifier) Then    ' Range
        nickname = getColumn(IND_Nickname, identifier).text
        indAttr = getRowValues(identifier, indAttr)
    Else  ' ニックネーム
        nickname = identifier
        indAttr = seachAndGetValues( _
                identifier, IND_Nickname, shIndividual, indAttr)
    End If
    '   入力途中かチェック
    If indAttr(3) = 0 Or IsEmpty(indAttr(1)) Or IsEmpty(indAttr(2)) Or IsEmpty(indAttr(3)) Then
        Exit Function
    End If
    If prediction Then PL = indAttr(5) Else PL = indAttr(4)
    Call clearMonster(mon, nickname, indAttr(0), PL, _
                indAttr(1), indAttr(2), indAttr(3))
    getIndividual = True
End Function

'   個体のわざの取得
Public Sub setIndividualAttacks(ByRef mon As Monster, _
            ByVal mode As Integer, _
            Optional ByVal prediction As Integer = 0, _
            Optional ByVal cel As Range = Nothing, _
            Optional ByVal atkDelay As Double = 0, _
            Optional ByVal considerEffect As Boolean = False)
    Dim atkNames As Variant
    Dim natks, satks, nflags, sflags As Variant
    '   わざ名の取得
    atkNames = Array( _
        IND_NormalAtk, IND_SpecialAtk1, IND_SpecialAtk2, _
        IND_TargetNormalAtk, IND_TargetSpecialAtk)
    If Not cel Is Nothing Then    ' セルで取得
        atkNames = getRowValues(cel, atkNames)
    Else  ' ニックネームで検索
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

'   パラメータを指定して個体を取得
Public Sub getMonster(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal PL As Double = 40, _
                    Optional ByVal indATK As Long = 15, _
                    Optional ByVal indDEF As Long = 15, _
                    Optional ByVal indHP As Long = 15, _
                    Optional ByVal defHP As Long = 0)
    '   デフォルト値
    If species = "" Then
        Call clearMonster(mon, , , PL, indATK, indDEF, indHP)
        Call dummySetPower(mon)
        mon.fullHP = Fix(mon.hpPower)
        mon.curHP = mon.fullHP
        Exit Sub
    End If
    Call clearMonster(mon, "", species, PL, indATK, indDEF, indHP, defHP)
End Sub

'   パラメータを指定して個体を取得
Public Sub getMonsterByPower(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal atk As Double = 100, _
                    Optional ByVal def As Double = 100, _
                    Optional ByVal hp As Double = 100)
    Dim attr, pow, limCPM(3) As Variant
    Dim CPM, max As Double
    Dim i As Integer
    Call clearMonster(mon, "", species)
    mon.atkPower = atk
    mon.defPower = def
    mon.hpPower = hp
    mon.fullHP = Fix(hp)
    mon.curHP = mon.fullHP
    mon.CP = Fix(atk * Sqr(def) * Sqr(hp) / 10)
    limCPM(3) = Array(-1E+100, 1E+100)  '   最も狭い範囲
    pow = Array(atk, def, hp)
    If species <> "" Then
        attr = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
        For i = 0 To 2
            limCPM(i) = Array(pow(i) / (attr(i) + 15), pow(i) / attr(i))
            If limCPM(3)(0) < limCPM(i)(0) Then limCPM(3)(0) = limCPM(i)(0)
            If limCPM(3)(1) > limCPM(i)(1) Then limCPM(3)(1) = limCPM(i)(1)
        Next
        If limCPM(3)(0) <= limCPM(3)(1) Then
            CPM = limCPM(3)(0)
        Else
            CPM = limCPM(3)(1)
        End If
        mon.PL = getPLbyCPM(CPM, False)
        mon.indATK = atk / CPM - attr(0)
        mon.indDEF = def / CPM - attr(1)
        mon.indHP = hp / CPM - attr(2)
    Else
        Call dummySetPlInd(mon)
    End If
End Sub

'   ダミーにおいて、パワーよりPLと個体値を適当に設定
Private Sub dummySetPlInd(ByRef mon As Monster)
    Dim CPM, max As Double
    CPM = Sqr(mon.CP + 0.5) / Sqr(4432) * 0.7903
    mon.PL = getPLbyCPM(CPM)
    max = mon.atkPower
    If max < mon.defPower Then max = mon.defPower
    If max < mon.hpPower Then max = mon.hpPower
    If max < 1 Then max = 1
    mon.indATK = mon.atkPower * 15 / max
    mon.indDEF = mon.defPower * 15 / max
    mon.indHP = mon.hpPower * 15 / max
End Sub

'   ダミーにおいて、適当なPLと個体値よりパワーを適当に設定
Private Sub dummySetPower(ByRef mon As Monster)
    Dim CPM, k(1), CP As Double
    CP = (getCPM(mon.PL) * Sqr(4432) / 0.7903) ^ 2
    mon.CP = Fix(CP)
    If mon.indATK < 1 Then mon.indATK = 1
    If mon.indDEF < 1 Then mon.indDEF = 1
    If mon.indHP < 1 Then mon.indHP = 1
    k(0) = mon.indDEF / mon.indATK
    k(1) = mon.indHP / mon.indATK
    mon.atkPower = Sqr(10 * CP / Sqr(k(0) * k(1)))
    mon.defPower = mon.atkPower * k(0)
    mon.hpPower = mon.atkPower * k(1)
End Sub

'   パラメータを指定して個体を取得
Public Sub getMonsterByCpHp(ByRef mon As Monster, _
                    Optional ByVal species As String = "", _
                    Optional ByVal CP As Long = 1000, _
                    Optional ByVal hp As Long = 100)
    Dim attr As Variant
    Dim CPd, HPd, CPM, ADmax, AD, CPpHP As Double
    Dim a, b, c, p, q, u, v, ind As Double
    Call clearMonster(mon, "", species)
    mon.CP = CP: mon.fullHP = hp: mon.curHP = hp
    CPd = CP + 0.5: HPd = hp + 0.5
    mon.hpPower = HPd
    CPpHP = 10 * CPd / Sqr(HPd) '
    If species <> "" Then
        attr = getSpcAttrs(species, Array("ATK", "DEF", "HP"))
        ADmax = (attr(0) + 15) * Sqr(attr(1) + 15)
        '   PLとADの決定
        CPM = getCPM(40)
        AD = CPpHP / CPM ^ 1.5 '   A√D40
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
        u = WorksheetFunction.Power(-q + Sqr(q ^ 2 + p ^ 2), 1 / 3)
        v = WorksheetFunction.Power(-q - Sqr(q ^ 2 + p ^ 2), 1 / 3)
        ind = u + v - a / 3
        mon.indATK = ind: mon.atkPower = (attr(0) + ind) * CPM
        mon.indDEF = ind: mon.defPower = (attr(1) + ind) * CPM
        mon.indHP = HPd / CPM - attr(2): mon.hpPower = HPd
    Else
        mon.atkPower = CPpHP ^ (2 / 3)
        mon.defPower = mon.atkPower
        Call dummySetPlInd(mon)
        Debug.Print "getMonsterByCpHp CP" & mon.CP & ", " _
                    & (mon.atkPower * Sqr(mon.defPower) * Sqr(mon.fullHP) / 10)
    End If
End Sub

'   わざの設定
Public Sub setAttacks(ByVal mode As Integer, ByRef mon As Monster, _
            Optional ByVal NormalAtk As Variant = "", _
            Optional ByVal SpecialAtk As Variant = "", _
            Optional ByVal atkDelay As Double = 0, _
            Optional ByVal considerEffect As Boolean = False)
    mon.mode = mode
    mon.atkNum = 0
    If Not IsArray(NormalAtk) Then NormalAtk = Array(NormalAtk)
    If Not IsArray(SpecialAtk) Then SpecialAtk = Array(SpecialAtk)
    Call setNormalAttacks(mon, NormalAtk, atkDelay)
    Call setSpecialAttacks(mon, SpecialAtk, atkDelay, considerEffect)
End Sub

'   クリア
Public Sub clearMonster(ByRef mon As Monster, _
        Optional ByVal nickname As String = "", _
        Optional ByVal species As String = "", _
        Optional ByVal PL As Double = 40, _
        Optional ByVal indATK As Long = 15, _
        Optional ByVal indDEF As Long = 15, _
        Optional ByVal indHP As Long = 15, _
        Optional ByVal defHP As Long = 0)
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
    '   その他クリア
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

'   PL、個体値より各パワーの計算
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

'   ダメージをまとめて計算
Private Sub calcDamages( _
            ByRef offence As Monster, ByRef deffence As Monster, _
            Optional ByVal isAna As Boolean = True, _
            Optional ByVal weather As Integer = 0, _
            Optional ByVal ignoreFactor As Boolean = False)
    Dim effect As Double
    Dim atk As Variant
    
    '   特殊効果計算。攻撃側
    With offence.attacks(offence.atkIndex(1).selected)
        effect = 1#
        If .isEffect Then
            effect = .effect.expect(IDX_SelfAtk) _
                    / .effect.expect(IDX_EnemyDef)
        End If
    End With
    '   特殊効果計算。防御側
    If deffence.atkIndex(1).selected >= 0 Then
        With deffence.attacks(deffence.atkIndex(1).selected)
            If .isEffect Then
                effect = effect * .effect.expect(IDX_EnemyAtk) _
                        / .effect.expect(IDX_SelfDef)
            End If
        End With
    End If
    '   解析がfalseでも、特殊効果があればtrueにして小数点以下をだす
    If isAna = False And effect <> 1 Then
        isAna = True
    End If
    Call calcADamage(offence.atkIndex(0).selected, _
            offence, deffence, isAna, weather, effect, ignoreFactor)
    Call calcADamage(offence.atkIndex(1).selected, _
            offence, deffence, isAna, weather, effect, ignoreFactor)
End Sub

'   ダメージを計算
Public Function calcADamage(ByRef atkIdx As Integer, _
            ByRef offence As Monster, ByRef deffence As Monster, _
            Optional ByVal isAna As Boolean = True, _
            Optional ByVal weather As Integer = 0, _
            Optional ByVal effect As Double = 1#, _
            Optional ByVal ignoreFactor As Boolean = False) As Double
    Dim fctr As Double
    If atkIdx < 0 Then Exit Function
    With offence.attacks(atkIdx)
        If deffence.defPower = 0 Then
            Debug.Print "defPower is 0. " & deffence.nickname
            .damage = 0
            Exit Function
        End If
        If ignoreFactor Then
            fctr = getFactor(offence.mode, offence.itype, _
                    .itype, Array(0, 0), 0)
        Else
            fctr = getFactor(offence.mode, offence.itype, _
                    .itype, deffence.itype, weather)
            fctr = fctr * effect
        End If
        .damage = (offence.atkPower _
                / deffence.defPower * .Power * fctr * 0.5) + 1
        If Not isAna Then .damage = Fix(.damage)
        calcADamage = .damage
    End With
End Function

'   係数の取得
Public Function getFactor(ByVal mode As Integer, _
                ByVal selfTypes As Variant, _
                ByVal atkType As Integer, _
                ByVal enemyTypes As Variant, _
                ByVal weather As Integer) As Double
    If Not IsArray(InterTypeFactorTable) Then Call makeInfluenceCache
    getFactor = 1#
    '   タイプ一致
    If atkType > 0 Then
        If atkType = selfTypes(0) Or atkType = selfTypes(1) Then
            getFactor = TypeMatchFactorValue
        End If
    End If
    '   タイプ相関
    getFactor = getFactor * InterTypeFactorTable(atkType, enemyTypes(0)) _
                * InterTypeFactorTable(atkType, enemyTypes(1))
    '   その他
    If mode = C_IdGym Then   ' ジムバトルでは天候ブースト
        getFactor = getFactor * WeatherFactorTable(atkType, weather)
    ElseIf mode = C_IdMtc Then   ' トレーナーバトルは1.3倍
        getFactor = getFactor * MatchBattleFactorValue
    End If
End Function

'   相関シートなどの表のキャッシュの作成
Public Sub makeInfluenceCache()
    InterTypeFactorTable = makeInterTypeTable
    WeatherFactorTable = makeWeatherFactorTable
    TypeMatchFactorValue = Range(R_TypeMatchFactor).value
    MatchBattleFactorValue = Range(R_MtcBtlFactor).value
    ChargeByDamage = Range(R_ChargeByDamage).value
End Sub

'   相関表の作成
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
            mark = Range(R_InterTypeInflu).cells(i, j).text
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

'   天候ブースト係数のキャッシュを作る
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
            boost = Range(R_WeatherBoost).cells(ti, 1).text
            wi = getWeatherIndex(boost)
            tbl(ti, wi) = Range(R_WeatherFactor).value
        End If
    Next
    makeWeatherFactorTable = tbl
End Function

'   タイプ一致係数
Private Function atkTypeFactor(ByVal self As Variant, ByVal atkType As Variant) As Double
    atkTypeFactor = 1
    Call type2idx(atkType)
    If atkType <> 0 Then
        Call types2idx(self)
        If self(0) = atkType Or self(1) = atkType Then
            atkTypeFactor = Range(R_TypeMatchFactor).value
        End If
    End If
End Function

'   相手との相性の係数
Public Function interTypeFactor(ByVal atkType As Integer, _
                ByVal enemy As Variant) As Double
    If Not IsArray(InterTypeFactorTable) Then Call makeInfluenceCache
    interTypeFactor = InterTypeFactorTable(atkType, enemy(0)) * InterTypeFactorTable(atkType, enemy(1))
End Function

'   天候ブーストの係数
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


'   KTまたはKTRによるランクを得る
'   返り値の配列の添字は0から。一要素は配列で、その各要素は以下。
'   0:KTR, 1:KT, 2:ニックネーム, 3:通常技,
'   4:通常技tDPS, 5:ゲージ技, 6:ゲージ技tDPS, 7:cDPS, 8:Cycle
Public Function getKtRank(ByVal rankNum As Long, _
        ByRef enemy As Monster, _
        Optional ByVal isKTR As Boolean = True, _
        Optional ByVal CPlimit As Variant = 0, _
        Optional ByVal prediction As Boolean = False, _
        Optional ByVal weather As Integer = 0, _
        Optional ByVal selfAtkDelay As Double = 0) As KtRank
    Dim rank() As SimInfo
    Dim cel As Range
    Dim ri As Long
    Dim predict As Integer
    Dim CP As Long
    Dim self As Monster
    Dim spesifiedEnemyAtk As Boolean
    Dim vtmp As SimInfo
    Dim cdpss As CDpsSet
    
    If Not IsNumeric(weather) Then weather = getWeatherIndex(weather)
    If Not IsArray(CPlimit) Then CPlimit = Array(0, 0)
    If prediction Then predict = 2 Else predict = 0
    With enemy  '   敵の技が特定されているか
        spesifiedEnemyAtk = _
                (.atkIndex(0).lower = .atkIndex(0).upper) _
                And (.atkIndex(1).lower = .atkIndex(1).upper)
    End With
    ReDim rank(rankNum)
    For Each cel In shIndividual.ListObjects(1).ListColumns(IND_CP).DataBodyRange
        dspProgress
        '   CPのチェック。下限・上限が有効でCPがその範囲を超えていたら次へ
        CP = cel.value
        If (CPlimit(1) > 0 And CP < CPlimit(1)) Or (CPlimit(0) > 0 And CPlimit(0) < CP) Then GoTo Continue
        '   個体の設定
        If Not getIndividual(cel, self, prediction) Then GoTo Continue
        Call setIndividualAttacks(self, enemy.mode, predict, cel, selfAtkDelay, _
                                    (enemy.mode = C_IdMtc))
        If self.PL = 0 Or self.atkIndex(0).lower < 0 Or self.atkIndex(1).lower < 0 Then GoTo Continue
        '   手持ち技の中で最も効果的な組み合わせによる結果を得る
        rank(rankNum) = getMinSimInfo(self, enemy, weather, _
                                spesifiedEnemyAtk, isKTR)
        '   ランクの更新
        ri = rankNum - 1
        Do While ri >= 0
            If rank(ri).KT > 0 Then
                If (isKTR And rank(ri).KTR <= rank(ri + 1).KTR) _
                Or (Not isKTR And rank(ri).KT <= rank(ri + 1).KT) Then
                    Exit Do
                End If
            End If
            vtmp = rank(ri): rank(ri) = rank(ri + 1): rank(ri + 1) = vtmp
            ri = ri - 1
        Loop
        If ri < rankNum - 1 Then
            ri = ri + 1
            Call resetHpAndEffectTrans(self)
            '   天候なし、敵タイプなしの場合のcDPS
            cdpss = calcCDPS(self, enemy, False, weather, True)
            With rank(ri)
                .cDPS = cdpss.cDPS
                .Cycle = cdpss.Cycle
            End With
            '   calcCDPS内で再計算された、効果無視のダメージにてtDPS計算
            With self.attacks(rank(ri).Attack(0).idx)
                rank(ri).Attack(0).name = .name
                rank(ri).Attack(0).tDPS = .damage / .idleTime
            End With
            With self.attacks(rank(ri).Attack(1).idx)
                rank(ri).Attack(1).name = .name
                rank(ri).Attack(1).tDPS = .damage / .idleTime
            End With
        End If
Continue:
    Next
    ReDim Preserve rank(rankNum - 1)
    getKtRank.rank = rank
End Function

'   手持ち技の中で最も効果的な組み合わせによる結果を得る
Private Function getMinSimInfo(ByRef self As Monster, _
                ByRef enemy As Monster, _
                ByVal weather As Integer, _
                ByVal spesifiedEnemyAtk As Boolean, _
                ByVal isKTR As Boolean) As SimInfo
    Dim ktrs As Variant
    Dim simi As SimInfo
    Dim nai, sai As Long
    
    simi.KTR = 1E+107: simi.KT = simi.KTR
    '   個体の通常・ゲージ技で、最小のKTRまたはKTのものを得る
    For nai = self.atkIndex(0).lower To self.atkIndex(0).upper
        self.atkIndex(0).selected = nai
        For sai = self.atkIndex(1).lower To self.atkIndex(1).upper
            self.atkIndex(1).selected = sai
            '   敵の技が特定されていたらKTRを取得
            If spesifiedEnemyAtk Then
                Call resetHpAndEffectTrans(self)
                Call resetHpAndEffectTrans(enemy)
                ktrs = simBattle(self, enemy, weather, True)
            Else    '   特定されていないので平均を得る
                ktrs = getAveKTR(self, enemy, weather)
            End If
            '   最小をsimiに記録
            If (isKTR And simi.KTR > ktrs(0)) _
                    Or (Not isKTR And simi.KT > ktrs(1)) Then
                With simi
                    .KTR = ktrs(0)
                    .KT = ktrs(1)
                    .Attack(0).idx = nai
                    .Attack(1).idx = sai
                End With
                If spesifiedEnemyAtk Then
                    With self
                        simi.Attack(0).stDPS = .given(0).damage / .given(0).time
                        If .given(1).time > 0 Then
                            simi.Attack(1).stDPS = .given(1).damage / .given(1).time
                        Else
                            simi.Attack(1).stDPS = 0
                        End If
                        simi.scDPS = (.given(0).damage + .given(1).damage) _
                                / (.given(0).time + .given(1).time)
                    End With
                Else
                    simi.Attack(0).stDPS = ktrs(2)
                    simi.Attack(1).stDPS = ktrs(3)
                    simi.scDPS = ktrs(4)
                End If
            End If
        Next
    Next
    '   同個体でKTR(KT)が最小のものの整形
    simi.nickname = self.nickname
    simi.PL = self.PL
'    With self.attacks(simi.Attack(0).idx)
'        simi.Attack(0).name = .name
'        simi.Attack(0).tDPS = .damage / .idleTime
'    End With
'    With self.attacks(simi.Attack(1).idx)
'        simi.Attack(1).name = .name
'        simi.Attack(1).tDPS = .damage / .idleTime
'    End With
    getMinSimInfo = simi
End Function

'   敵の技を総当りして、KT、KTRの平均値を取得。
'   戻り値は配列。その要素は配列で
'   0:KRT, 1:自KT, 2:通常わざtDPS, 3:ゲージわざtDPS, 4:cDPS
Public Function getAveKTR( _
        ByRef self As Monster, ByRef enemy As Monster, _
        Optional ByVal weather As String = "") As Variant
    Dim sumKtr(4) As Variant
    Dim ktrs As Variant
    Dim nai, sai, cnt, i As Long
    '   クリア
    For i = 0 To UBound(sumKtr)
        sumKtr(i) = 0#
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
                    sumKtr(2) = sumKtr(2) _
                        + (.given(0).damage / .given(0).time)
                    If .given(1).time > 0 Then
                        sumKtr(3) = sumKtr(3) _
                            + (.given(1).damage / .given(1).time)
                    End If
                    sumKtr(4) = sumKtr(4) _
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

'   ターゲットCPに合う個体値の探索
Function getFitIndiv(ByRef self As Monster, ByVal tCP As Long, _
                    Optional ByVal rnkNum As Integer = 100) As Variant
    Dim atk, def, hp, cnt As Long
    Dim PLlim(1), CPM(), CPM2(), DH, ADH, drv, drvMax, CP, cpg As Double
    Dim n, i As Integer
    Dim sv, rnk(), vtmp As Variant
    Dim enemy As Monster
    
    '   charge countを計算しておく
    If Not calcChargeCount(self) Then Exit Function
    Call getMonsterByPower(enemy)
    sv = getSpcAttrs(self.species, Array("ATK", "DEF", "HP"))
    '   種族値よりPLの範囲を得、CMP^2を計算しておく
    cpg = sv(0) * Sqr(sv(1) * sv(2)) / 10
    PLlim(1) = getPLbyCpg(tCP, cpg)
    cpg = (sv(0) + 15) * Sqr((sv(1) + 15) * (sv(2) + 15)) / 10
    PLlim(0) = getPLbyCpg(tCP, cpg)
    n = (PLlim(1) - PLlim(0)) * 2
    ReDim CPM(n), CPM2(n)
    For i = 0 To n
        CPM(i) = getCPM(PLlim(0) + 0.5 * i)
        CPM2(i) = CPM(i) ^ 2
    Next
    ReDim rnk(rnkNum)
    For def = 0 To 15
        self.indDEF = def
        For hp = 0 To 15
            self.indHP = hp
            DH = Sqr((sv(1) + def) * (sv(2) + hp))
            For atk = 0 To 15
                self.indATK = atk
                ADH = (sv(0) + atk) * DH / 10
                drvMax = 0
                For i = n To 0 Step -1
                    CP = CPM2(i) * ADH
                    If CP < tCP - 200 Then Exit For
                    If CP < tCP + 1 Then
                        self.atkPower = (self.indATK + sv(0)) * CPM(i)
                        self.defPower = (self.indDEF + sv(1)) * CPM(i)
                        self.hpPower = (self.indHP + sv(2)) * CPM(i)
                        drvMax = calcTCP(self, enemy)
                        Exit For
'                        drv = calcTCP(self, enemy)
'                        If drvMax < drv Then
'                            drvMax = drv
'                        End If
                    End If
                Next
                If drvMax > 0 Then
                    cnt = cnt + 1
                    rnk(rnkNum) = Array(atk, def, hp, drvMax)
                    For i = rnkNum - 1 To 0 Step -1
                        If IsArray(rnk(i)) Then
                            If rnk(i)(3) > rnk(i + 1)(3) Then Exit For
                        End If
                        vtmp = rnk(i): rnk(i) = rnk(i + 1): rnk(i + 1) = vtmp
                    Next
                End If
            Next
        Next
    Next
    If cnt > rnkNum Then cnt = rnkNum
    If cnt > 0 Then
        ReDim Preserve rnk(cnt - 1)
        getFitIndiv = rnk
    End If
End Function

'   TCPの算出
Public Function calcTCP(ByRef self As Monster, ByRef enemy As Monster) As Double
    Dim damage, period As Double
    Dim atkIdx As Variant
    
    Call calcDamages(self, enemy, True)
    atkIdx = getAttackIndex(self)
    With self
        damage = .attacks(atkIdx(0)).damage * self.chargeCount + .attacks(atkIdx(1)).damage
        period = .attacks(atkIdx(0)).idleTime * self.chargeCount + .attacks(atkIdx(1)).idleTime
    End With
    calcTCP = damage / period * Fix(self.hpPower) / (1000 / self.defPower + 1)
End Function




